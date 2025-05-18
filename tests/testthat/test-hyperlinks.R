href_list <- c(
  "https://github.com/davidgohel",
  "http://www.google.com?a=1&b=2",
  "http://www.google.com/folder1",
  "http://leginfo.legislature.ca.gov/faces/codes_displaySection.xhtml?lawCode=EDC&sectionNum=52164.1",
  "https://spec.tass.ru/avangard-na-shabolovke/?from=newsapp&utm_partner_id=149"
)
block_links <- lapply(href_list, function(link) {
  fpar(hyperlink_ftext(text = "link", href = link))
})
block_links <- do.call(block_list, block_links)

test_that("add hyperlink in docx", {
  x <- read_docx()
  x <- body_add_blocks(x, block_links)
  filename <- print(x, target = tempfile(fileext = ".docx"))

  x <- read_docx(path = filename)
  rel_df <- x$doc_obj$rel_df()
  expect_true(all(href_list %in% rel_df$target))
  subset_rel <- rel_df[rel_df$target %in% href_list, ]
  if (nrow(subset_rel) > 0) {
    expect_true(all(subset_rel$target %in% href_list))
    expect_true(all(subset_rel$target_mode %in% "External"))
    expect_match(subset_rel$type, "^http://schemas(.*)hyperlink$")

    body <- docx_body_xml(x)
    nodes_hyperlink <- xml_find_all(body, "//w:hyperlink")
    expect_length(nodes_hyperlink, 5)
    expect_true(all(xml_attr(nodes_hyperlink, "id") %in% subset_rel$id))
  }
})

test_that("add hyperlink in pptx", {
  x <- read_pptx()
  x <- add_slide(x, "Title and Content")
  x <- ph_with(x, block_links, location = ph_location_type())
  filename <- print(x, target = tempfile(fileext = ".pptx"))

  x <- read_pptx(path = filename)
  slide <- x$slide$get_slide(x$cursor)

  rel_df <- slide$rel_df()
  expect_true(all(href_list %in% rel_df$target))
  subset_rel <- rel_df[rel_df$target %in% href_list, ]
  if (nrow(subset_rel) > 0) {
    expect_true(all(subset_rel$target %in% href_list))
    expect_true(all(subset_rel$target_mode %in% "External"))
    expect_match(subset_rel$type, "^http://schemas(.*)hyperlink$")

    nodes_hlinkclick <- xml_find_all(
      slide$get(),
      "/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick"
    )
    expect_length(nodes_hlinkclick, 5)
    expect_true(all(xml_attr(nodes_hlinkclick, "id") %in% subset_rel$id))
  }
})
