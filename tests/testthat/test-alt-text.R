test_that("add alt text to ggplot", {
  skip_if_not_installed("ggplot2")
  require(ggplot2)

  alt_text <- "Alt text added with 'alt_text='."

  gg <- ggplot(mtcars, aes(factor(cyl))) +
    geom_bar()

  doc <- read_pptx()
  doc <- add_slide(doc)
  doc <- ph_with(doc,
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
  doc <- add_slide(doc)
  doc <- ph_with(doc,
    value = gg  + labs(alt = alt_labs),
    location = ph_location_type("body")
  )
  xmldoc <- doc$slide$get_slide(id = 2)$get()
  expect_equal(
    xml_attr(xml_find_all(xmldoc, "//p:nvPicPr//p:cNvPr"), "descr"),
    alt_labs
  )

  doc <- add_slide(doc)
  doc <- ph_with(doc,
                 value = gg,
                 location = ph_location_type("body")
  )
  xmldoc <- doc$slide$get_slide(id = 3)$get()
  expect_equal(
    xml_attr(xml_find_all(xmldoc, "//p:nvPicPr//p:cNvPr"), "descr"),
    ""
  )
})
