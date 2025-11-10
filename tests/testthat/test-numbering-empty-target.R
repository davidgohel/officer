test_that("body_import_docx preserves list formatting from officedown documents", {
  # This test verifies the fix for issue #694
  # When importing from officedown-generated documents with high numIds (1000+),
  # the function should properly add all numbering definitions

  skip_if_not_installed("officedown")
  skip_if_not_installed("rmarkdown")

  # Create RMarkdown file with both ordered and bulleted lists
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines("---
output: officedown::rdocx_document
---

# Lists Test

Ordered list:

1. Item 1
2. Item 2

Bulleted list:

- Item A
- Item B
", rmd_file)

  output_file <- tempfile(fileext = ".docx")
  rmarkdown::render(rmd_file, quiet = TRUE, output_file = output_file)

  target <- read_docx()
  target <- body_add_par(target, "Frontmatter", style = "heading 1")

  target <- body_import_docx(target, output_file, par_style_mapping = list(
    "Normal" = c("First Paragraph", "Compact")
  ))

  target_file <- tempfile(fileext = ".docx")
  print(target, target = target_file)

  result <- read_docx(target_file)
  result_summary <- docx_summary(result)
  list_items <- result_summary[!is.na(result_summary$num_id), ]

  expect_equal(nrow(list_items), 4)

  # Verify numbering definitions exist and have correct formats
  library(xml2)
  result_numbering <- read_xml(file.path(result$package_dir, "word/numbering.xml"))

  # Check that the numIds used in the document have corresponding definitions
  for (num_id in unique(list_items$num_id)) {
    num_node <- xml_find_first(result_numbering, sprintf("//w:num[@w:numId='%d']", num_id))
    expect_false(is.na(num_node), info = sprintf("num %d should exist in numbering.xml", num_id))

    # Verify it has a valid abstractNum reference
    abstract_ref <- xml_find_first(num_node, ".//w:abstractNumId")
    expect_false(is.na(abstract_ref))
    abstract_id <- xml_attr(abstract_ref, "val")

    abstract_node <- xml_find_first(result_numbering, sprintf("//w:abstractNum[@w:abstractNumId='%s']", abstract_id))
    expect_false(is.na(abstract_node), info = sprintf("abstractNum %s should exist", abstract_id))
  }
})
