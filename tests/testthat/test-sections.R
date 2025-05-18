txt_lorem <- "Tellus nunc facilisis, quam tempor tempor. Purus lectus eros metus turpis mattis platea praesent sed. Sed ac, interdum tincidunt auctor mi tempor a pulvinar ut massa. Accumsan id fusce at platea! Magna tincidunt parturient ultricies netus vestibulum libero ex ornare in molestie, risus. Tempor nibh mollis pharetra vitae mauris nec cum et curabitur, lorem quam semper, lectus. Rutrum in in at eu non, diam. Eu fermentum ac donec eleifend. Condimentum neque etiam sit euismod mollis sollicitudin dictumst."
img.file <- file.path(R.home("doc"), "html", "logo.jpg")

bl_header_first <- block_list(
  fpar(
    ftext("header_first "),
    external_img(src = img.file, height = 1.06 / 2, width = 1.39 / 2)
  )
)

bl_header_even <- block_list(fpar(ftext("header_even")))
bl_header_odd <- block_list(fpar(ftext("header_odd")))
bl_footer_first <- block_list(fpar(ftext("footer_first")))
bl_footer_even <- block_list(fpar(ftext("footer_even")))
bl_footer_odd <- block_list(fpar(ftext("footer_odd")))

x <- read_docx()

for (i in 1:40) {
  x <- body_add_par(x, value = txt_lorem)
}


x <- body_set_default_section(
  x,
  value = prop_section(
    page_margins = page_mar(bottom = 2),
    page_size = page_size(orient = "portrait"),
    header_first = bl_header_first,
    footer_first = bl_footer_first,
    header_default = bl_header_odd,
    header_even = bl_header_even,
    footer_default = bl_footer_odd,
    footer_even = bl_footer_even
  )
)
filename <- print(x, target = tempfile(fileext = ".docx"))

test_that("add header and footer in docx", {
  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)header$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 3)
  if (nrow(subset_rel) > 0) {
    pattern <- "^(header)([0-9]+)(\\.xml)$"
    expect_true(all(grepl(pattern, subset_rel$target)))
    nodes_def_sectpr <- xml_find_all(
      docx_body_xml(x),
      "//w:body/w:sectPr/w:headerReference"
    )
    expect_length(nodes_def_sectpr, 3)
    nodes_def_sectpr <- xml_find_all(
      docx_body_xml(x),
      "//w:body/w:sectPr/w:footerReference"
    )
    expect_length(nodes_def_sectpr, 3)

    first_str <- readLines(file.path(
      x$package_dir,
      "word/_rels/header3.xml.rels"
    ))
    expect_true(any(grepl("\\.jpg", first_str)))
  }
  nodes_first_sectpr <- xml_find_first(
    docx_body_xml(x),
    "//w:body/w:sectPr/w:headerReference[@w:type='first']"
  )
  expect_false(inherits(nodes_first_sectpr, "xml_missing"))

  id <- xml_attr(nodes_first_sectpr, "id")
  target <- subset_rel$target[subset_rel$id %in% id]

  xml_doc <- read_xml(file.path(x$package_dir, "word", target))
  expect_equal(xml_text(xml_doc), "header_first ")

  sect_title_node <- xml_find_first(
    docx_body_xml(x),
    "//w:body/w:sectPr/w:titlePg"
  )
  expect_false(inherits(sect_title_node, "xml_missing"))
})


library(officer)
txt_lorem <- "Tellus nunc facilisis, quam tempor tempor. Purus lectus eros metus turpis mattis platea praesent sed. Sed ac, interdum tincidunt auctor mi tempor a pulvinar ut massa. Accumsan id fusce at platea! Magna tincidunt parturient ultricies netus vestibulum libero ex ornare in molestie, risus. Tempor nibh mollis pharetra vitae mauris nec cum et curabitur, lorem quam semper, lectus. Rutrum in in at eu non, diam. Eu fermentum ac donec eleifend. Condimentum neque etiam sit euismod mollis sollicitudin dictumst."

x <- read_docx()

for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}

x <- body_end_block_section(
  x,
  value = block_section(prop_section(
    page_size = page_size(orient = "portrait")
  ))
)

for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}

x <- body_end_block_section(
  x,
  value = block_section(prop_section(
    page_size = page_size(orient = "landscape")
  ))
)

for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}

x <- body_set_default_section(
  x,
  value = prop_section(
    page_margins = page_mar(),
    page_size = page_size(orient = "portrait")
  )
)
filename <- print(x, target = tempfile(fileext = ".docx"))


test_that("add docx landscape and portrait page dims", {
  x <- read_docx(path = filename)
  node_types <- xml_find_all(
    docx_body_xml(x),
    "//w:sectPr/w:pgSz"
  )
  types <- xml_attr(node_types, "orient")
  expect_equal(types, c("portrait", "landscape", "portrait"))

  h <- xml_attr(node_types, "h")
  expect_equal(h, c("16838", "11906", "16838"))
  w <- xml_attr(node_types, "w")
  expect_equal(w, c("11906", "16838", "11906"))
})
