library(doconv)
library(officer)
library(minimage)
source("inst/examples/example_read_docx_1.R")
to_miniature(comprehensive_file, fileout = "man/figures/read_docx.png", row = c(1, 1))
source("inst/examples/example_ph_with.R")
to_miniature(
  filename = fileout,
  fileout = "man/figures/ph_with_doc_1.png",
  width = 800,
  row = c(1, 1, 2, 2, 3, 3, 4)
)
source("inst/examples/example_body_add_1.R")
to_miniature(output_file, fileout = "man/figures/body_add_doc_1.png", dpi = 150, width = 850, row = c(1, 1, 2, 3))
source("inst/examples/example_body_set_default_section.R")
to_miniature(output_file, fileout = "man/figures/body_set_default_section_doc_1.png", dpi = 150)
source("inst/examples/example_body_end_block_section.R")
to_miniature(output_file_1, fileout = "man/figures/body_end_block_section_doc_1.png", dpi = 150)
source("inst/examples/example_prop_section.R")
to_miniature(output_file_1, fileout = "man/figures/prop_section_doc_1.png", dpi = 150, width = 850)

minimage::compress_images(input = "man/figures", output = "man/figures", overwrite = TRUE)

