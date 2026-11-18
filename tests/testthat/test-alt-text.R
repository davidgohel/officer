test_that("add alt text to ggplot", {
  skip_if_not_installed("ggplot2")
  require(ggplot2)

  alt_text <- "Alt text added with 'alt_text='."

  gg <- ggplot(mtcars, aes(factor(cyl))) +
    geom_bar()

  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    value = gg,
    location = ph_location_type("body"),
    alt_text = alt_text
  )

  xmldoc <- doc$slide$get_slide(id = 1)$get()
  expect_equal(
    xml_attr(xml_find_all(xmldoc, "//p:nvPicPr//p:cNvPr"), "descr"),
    alt_text
  )

  alt_labs <- "Alt text added with 'ggplot2::labs(alt=)'."
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    value = gg + labs(alt = alt_labs),
    location = ph_location_type("body")
  )
  xmldoc <- doc$slide$get_slide(id = 2)$get()
  expect_equal(
    xml_attr(xml_find_all(xmldoc, "//p:nvPicPr//p:cNvPr"), "descr"),
    alt_labs
  )

  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(doc, value = gg, location = ph_location_type("body"))
  xmldoc <- doc$slide$get_slide(id = 3)$get()
  expect_equal(
    xml_attr(xml_find_all(xmldoc, "//p:nvPicPr//p:cNvPr"), "descr"),
    ""
  )
})

test_that("alt text and frame names are escaped in the XML", {
  img <- file.path(R.home("doc"), "html", "logo.jpg")
  text <- "A & B < 5 \"x\""

  # pptx: alt text and name of an image, name of a text and of a table
  doc <- read_pptx()
  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    external_img(img, width = 1, height = 1, alt = text),
    location = ph_location_type("body", newlabel = text)
  )
  xmldoc <- doc$slide$get_slide(id = 1)$get()
  node <- xml_find_first(xmldoc, "//p:nvPicPr/p:cNvPr")
  expect_equal(xml_attr(node, "descr"), text)
  expect_equal(xml_attr(node, "name"), text)

  doc <- add_slide(doc, "Title and Content")
  doc <- ph_with(
    doc,
    "hello",
    location = ph_location_type("body", newlabel = text)
  )
  doc <- ph_with(
    doc,
    head(iris, 2),
    location = ph_location_type("body", newlabel = text)
  )
  xmldoc <- doc$slide$get_slide(id = 2)$get()
  frame_names <- xml_attr(xml_find_all(xmldoc, "//p:cNvPr"), "name")
  expect_equal(sum(frame_names %in% text), 2)

  # docx: inline and floating images
  for (value in list(
    external_img(img, width = 1, height = 1, alt = text),
    floating_external_img(img, width = 1, height = 1, alt = text)
  )) {
    doc <- body_add_fpar(read_docx(), fpar(value))
    node <- xml_find_first(doc$doc_obj$get(), "//wp:docPr")
    expect_equal(xml_attr(node, "descr"), text)
  }

  # html
  expect_match(
    to_html(external_img(img, width = 1, height = 1, alt = text)),
    "alt=\"A &amp; B &lt; 5 &quot;x&quot;\"",
    fixed = TRUE
  )
})
