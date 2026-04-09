test_that("settings works", {
  x <- read_docx()
  x <- docx_set_settings(
    x = x,
    zoom = 1,
    default_tab_stop = .5,
    hyphenation_zone = .25,
    decimal_symbol = ".",
    list_separator = ";",
    compatibility_mode = "15",
    even_and_odd_headers = TRUE,
    auto_hyphenation = FALSE
  )
  file <- print(x, target = tempfile(fileext = ".docx"))
  x <- read_docx(path = file)
  expect_equal(x$settings$zoom, 1)
  expect_equal(x$settings$list_separator, ";")
  expect_true(x$settings$even_and_odd_headers)
})

test_that("settings preserves existing XML elements", {
  template <- system.file(
    "doc_examples", "example.docx",
    package = "officer"
  )
  x <- read_docx(template)

  # read original settings.xml
  settings_before <- xml2::read_xml(
    file.path(x$package_dir, "word", "settings.xml")
  )
  tags_before <- xml2::xml_name(xml2::xml_children(settings_before))

  x <- docx_set_settings(x, zoom = 2)
  file <- print(x, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(file, unpack_dir)
  settings_after <- xml2::read_xml(
    file.path(unpack_dir, "word", "settings.xml")
  )
  tags_after <- xml2::xml_name(xml2::xml_children(settings_after))

  # all original tags should still be present
  expect_true(all(tags_before %in% tags_after))

  # zoom should be updated
  zoom_node <- xml2::xml_child(settings_after, "w:zoom")
  expect_equal(xml2::xml_attr(zoom_node, "percent"), "200")
})

test_that("docx_embed_font embeds font files", {
  gdtools::register_liberationsans()
  sysfonts <- gdtools::sys_fonts()
  libsans <- sysfonts[sysfonts$family == "Liberation Sans", ]

  x <- read_docx()
  x <- docx_embed_font(
    x,
    font_family = "Liberation Sans",
    regular = libsans$path[libsans$style == "Regular"],
    bold = libsans$path[libsans$style == "Bold"]
  )
  file <- print(x, target = tempfile(fileext = ".docx"))

  unpack_dir <- tempfile()
  unpack_folder(file, unpack_dir)

  # odttf files created
  odttfs <- list.files(
    file.path(unpack_dir, "word", "fonts"),
    pattern = "[.]odttf$"
  )
  expect_equal(length(odttfs), 2L)

  # fontTable.xml has embed entries
  ft <- xml2::read_xml(
    file.path(unpack_dir, "word", "fontTable.xml")
  )
  ns <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
  font_node <- xml2::xml_find_first(
    ft, "w:font[@w:name='Liberation Sans']", ns = ns
  )
  expect_false(inherits(font_node, "xml_missing"))
  embed_reg <- xml2::xml_child(font_node, "w:embedRegular")
  embed_bold <- xml2::xml_child(font_node, "w:embedBold")
  expect_false(inherits(embed_reg, "xml_missing"))
  expect_false(inherits(embed_bold, "xml_missing"))

  # fontTable.xml.rels has relationships
  rels <- xml2::read_xml(
    file.path(unpack_dir, "word", "_rels", "fontTable.xml.rels")
  )
  rel_nodes <- xml2::xml_find_all(rels, "d1:Relationship",
    ns = xml2::xml_ns(rels))
  expect_equal(length(rel_nodes), 2L)

  # settings.xml has embedTrueTypeFonts
  settings <- xml2::read_xml(
    file.path(unpack_dir, "word", "settings.xml")
  )
  embed_node <- xml2::xml_child(settings, "w:embedTrueTypeFonts")
  expect_false(inherits(embed_node, "xml_missing"))
})

test_that("docx_embed_font auto-detects font files", {
  gdtools::register_liberationsans()
  x <- read_docx()
  expect_message(
    x <- docx_embed_font(x, font_family = "Liberation Sans"),
    "Liberation Sans"
  )
  file <- print(x, target = tempfile(fileext = ".docx"))
  unpack_dir <- tempfile()
  unpack_folder(file, unpack_dir)

  odttfs <- list.files(
    file.path(unpack_dir, "word", "fonts"),
    pattern = "[.]odttf$"
  )
  expect_gte(length(odttfs), 1L)

  # silent mode
  x2 <- read_docx()
  expect_no_message(
    suppressMessages(
      docx_embed_font(x2, font_family = "Liberation Sans")
    )
  )
})

test_that("docx_embed_font auto-detect errors on unknown family", {
  x <- read_docx()
  expect_error(
    docx_embed_font(x, font_family = "NonExistentFont12345"),
    "not found"
  )
})

test_that("docx_embed_font validates inputs", {
  x <- read_docx()
  expect_error(
    docx_embed_font(x, "Test", regular = "nonexistent.ttf"),
    "does not exist"
  )
  expect_error(
    docx_embed_font("not_rdocx", "Test", regular = tempfile()),
    "rdocx"
  )
})
