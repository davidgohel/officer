context("add elements in slides")

source("utils.R")

test_that("add text into placeholder", {
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_empty(type = "body")
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)

  small_red <- fp_text(color = "red", font.size = 14)
  doc <- doc %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "chunk 1", style = small_red )
  sm <- slide_summary(doc)
  expect_equal(sm[1, ]$text, "chunk 1")

  doc <- doc %>%
    ph_add_text(str = "this is ", style = small_red, pos = "before" )
  sm <- slide_summary(doc)
  expect_equal(sm[1, ]$text, "this is chunk 1")
})

test_that("ph_add_par append when text alrady exists", {
  small_red <- fp_text(color = "red", font.size = 14)
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "This is a ") %>%
    ph_add_par(level = 2)
  doc <- doc %>%
    ph_add_text(str = "test", style = small_red )

  sm <- slide_summary(doc)
  expect_equal(sm$text, "This is a test")

  xmldoc <- doc$slide$get_slide(1)$get()
  txt <- xml_find_all(xmldoc, "//p:spTree/p:sp/p:txBody/a:p") %>% xml_text()
  expect_equal(txt, c("This is a ", "test"))
})


test_that("ph_add_text with hyperlink", {
  small_red <- fp_text(color = "red", font.size = 14)
  href_ <- "https://cran.r-project.org"
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "This is a ") %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "test", style = small_red, href = href_ )

  xmldoc <- doc$slide$get_slide(1)$get()
  rel_df <- doc$slide$get_slide(1)$rel_df()
  expect_true( href_ %in% rel_df$target )
  row_id_ <- which( rel_df$target_mode %in% "External" & rel_df$target %in% href_ )

  rid <- rel_df[row_id_, "id"]
  xpath_ <- sprintf("//p:sp/p:txBody/a:p/a:r[a:rPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_ )
  expect_equal( xml_text(node_), "test")
})

test_that("add img into placeholder", {
  skip_on_os("windows")
  img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_img(type = "body", src = img.file,
                height = 1.06, width = 1.39 )
  sm <- slide_summary(doc)

  expect_equal(nrow(sm), 1)
  expect_equal(sm$cx, 1.39)
  expect_equal(sm$cy, 1.06)


})

test_that("add formatted par into placeholder", {
  bold_face <- shortcuts$fp_bold(font.size = 30)
  bold_redface <- update(bold_face, color = "red")

  fpar_ <- fpar(ftext("Hello ", prop = bold_face),
                ftext("World", prop = bold_redface ),
                ftext(", how are you?", prop = bold_face ) )

  doc <- read_pptx() %>%
    add_slide(layout = "Title and Content", master = "Office Theme") %>%
    ph_empty(type = "body") %>%
    ph_add_fpar(value = fpar_, type = "body", level = 2)
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 1)
  expect_equal(sm[1, ]$text, "Hello World, how are you?")

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  cols <- xml_find_all(xmldoc, "//a:rPr/a:solidFill/a:srgbClr") %>% xml_attr("val")
  expect_equal(cols, c("000000", "FF0000", "000000") )
  expect_equal(xml_find_all(xmldoc, "//a:rPr") %>% xml_attr("b"), rep("1",3))
})


test_that("add xml into placeholder", {
  xml_str <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr/>\n<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Hello world 1</a:t></a:r></a:p></p:txBody></p:sp>"

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_from_xml(type = "body", value = xml_str) %>%
    ph_from_xml_at(left = 1, top = 1, width = 3, height = 3, value = xml_str)
  sm <- slide_summary(doc)
  expect_equal(nrow(sm), 2)
  expect_equal(sm[1,]$text, "Hello world 1")
  expect_equal(sm[2,]$text, "Hello world 1")
})

test_that("link to another shape", {
  small_red <- fp_text(color = "red", font.size = 14)

  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "This is a ") %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "test", style = small_red, slide_index = 1 )

  xmldoc <- doc$slide$get_slide(1)$get()
  rel_df <- doc$slide$get_slide(1)$rel_df()
  expect_true( "slide1.xml" %in% rel_df$target )
  row_id_ <- which( rel_df$target %in% "slide1.xml" )
  rid <- rel_df[row_id_, "id"]
  xpath_ <- sprintf("//p:sp/p:txBody/a:p/a:r[a:rPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_ )
  expect_equal( xml_text(node_), "test")
})

test_that("ph_add_text with hyperlink", {
  small_red <- fp_text(color = "red", font.size = 14)
  href_ <- "https://cran.r-project.org"
  doc <- read_pptx() %>%
    add_slide("Title and Content", "Office Theme") %>%
    ph_with_text(type = "body", str = "This is a ") %>%
    ph_add_par(level = 2) %>%
    ph_add_text(str = "test", style = small_red, href = href_ )

  xmldoc <- doc$slide$get_slide(1)$get()
  rel_df <- doc$slide$get_slide(1)$rel_df()
  expect_true( href_ %in% rel_df$target )
  row_id_ <- which( rel_df$target_mode %in% "External" & rel_df$target %in% href_ )

  rid <- rel_df[row_id_, "id"]
  xpath_ <- sprintf("//p:sp/p:txBody/a:p/a:r[a:rPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_ )
  expect_equal( xml_text(node_), "test")
})



test_that("slidelink shape", {

  doc <- read_pptx() %>%
    add_slide(layout = "Title and Content", master = "Office Theme") %>%
    ph_with_text(type = "title", str = "Un titre 1") %>%
    ph_with_text(type = "body", str = "text 1") %>%
    add_slide(layout = "Title and Content", master = "Office Theme") %>%
    ph_with_text(type = "title", str = "Un titre 2") %>%
    on_slide( index = 1)

  doc <- doc %>%
    ph_slidelink(type = "body", slide_index = 2 )

  rel_df <- doc$slide$get_slide(1)$rel_df()

  slide_filename <- doc$slide$get_metadata()$name[2]

  expect_true( slide_filename %in% rel_df$target )
  row_id_ <- which( is.na(rel_df$target_mode) & rel_df$target %in% slide_filename )

  rid <- rel_df[row_id_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_ )
  expect_false( inherits(node_, "xml_missing") )
})

test_that("hyperlink shape", {

  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with_text(x = doc, type = "title", str = "Un titre 1")
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with_text(x = doc, type = "title", str = "Un titre 2")
  doc <- on_slide(doc, 1)
  doc <- ph_hyperlink(x = doc, id_chr = "2", href = "https://cran.r-project.org")

  rel_df <- doc$slide$get_slide(1)$rel_df()

  expect_true( "https://cran.r-project.org" %in% rel_df$target )
  row_id_ <- which( !is.na(rel_df$target_mode) & rel_df$target %in% "https://cran.r-project.org" )

  rid <- rel_df[row_id_, "id"]
  xpath_ <- sprintf("//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick/@r:id='%s']", rid)
  node_ <- xml_find_first(doc$slide$get_slide(1)$get(), xpath_ )
  expect_false( inherits(node_, "xml_missing") )
})

unlink("*.pptx")

