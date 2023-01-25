txt_lorem <- "Tellus nunc facilisis, quam tempor tempor. Purus lectus eros metus turpis mattis platea praesent sed. Sed ac, interdum tincidunt auctor mi tempor a pulvinar ut massa. Accumsan id fusce at platea! Magna tincidunt parturient ultricies netus vestibulum libero ex ornare in molestie, risus. Tempor nibh mollis pharetra vitae mauris nec cum et curabitur, lorem quam semper, lectus. Rutrum in in at eu non, diam. Eu fermentum ac donec eleifend. Condimentum neque etiam sit euismod mollis sollicitudin dictumst."
img.file <- file.path( R.home("doc"), "html", "logo.jpg" )

value_first <- block_list(fpar(ftext("hello first "), external_img(src = img.file, height = 1.06/2, width = 1.39/2)))
value_even <- block_list(fpar(ftext("hello even")))
value_default <- block_list(fpar(ftext("hello default")))
footer_first <- block_list(fpar(ftext("bye first")))
footer_even <- block_list(fpar(ftext("bye even")))
footer_default <- block_list(fpar(ftext("bye default")))

ps <- prop_section(
  header_default = value_default, footer_default = footer_default,
  header_first = value_first, footer_first = footer_first,
  header_even = value_even, footer_even = footer_even
)
x <- read_docx()
for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}
x <- body_end_block_section(x, value = block_section(ps))
for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}
x <- body_set_default_section(
  x,
  value = prop_section(
    page_margins = page_mar(bottom = 2),
    page_size = page_size(width = 21 / 2.54, height = 29.7 / 2.54, orient = "portrait"),
    header_default = block_list(fpar(ftext("yok yok on top"))),
    header_even = block_list(fpar(ftext("yok yok on even top"))),
    footer_default = block_list(fpar(ftext("Chewbacca on bottom"))),
    footer_even = block_list(fpar(ftext("Chewbacca on even bottom")))
  )
)
filename <- print(x, target = tempfile(fileext = ".docx"))

test_that("add header and footer in docx", {
  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  subset_rel <- rel_df[grepl("^http://schemas(.*)header$", rel_df$type), ]
  expect_true(nrow(subset_rel) == 5)
  if (nrow(subset_rel) > 0) {
    pattern <- "^(header)([0-9]+)(\\.xml)$"
    expect_true(all(grepl(pattern, subset_rel$target)))
    nodes_def_sectpr <- xml_find_all(docx_body_xml(x), "//w:body/w:sectPr/w:headerReference")
    expect_length(nodes_def_sectpr, 2)
    nodes_def_sectpr <- xml_find_all(docx_body_xml(x), "//w:body/w:sectPr/w:footerReference")
    expect_length(nodes_def_sectpr, 2)

    first_str <- readLines(file.path(x$package_dir, "word/_rels/header5.xml.rels"))
    expect_true(any(grepl("\\.jpg", first_str)))
  }
})

test_that("margin and sizes are copied", {
  x <- read_docx(path = filename)
  node_def_sectpr <- xml_find_first(docx_body_xml(x), "//w:body/w:sectPr")
  nodes_sectpr <- xml_find_all(docx_body_xml(x), "//w:pPr/w:sectPr")
  pgMar <- as.character(xml_child(node_def_sectpr, "w:pgMar"))
  pgSz <- as.character(xml_child(node_def_sectpr, "w:pgSz"))
  for (i in seq_along(nodes_sectpr)) {
    expect_equal(as.character(xml_child(nodes_sectpr[[i]], "w:pgMar")), pgMar)
    expect_equal(as.character(xml_child(nodes_sectpr[[i]], "w:pgSz")), pgSz)
  }
})
