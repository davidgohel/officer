library(logger)
library(imgpress)
library(magrittr)

log_layout(layout_glue_colors)
log_threshold(TRACE)

man_names <- package_man_names(package_name = "officer")
drop <- c("fpar")
man_names <- setdiff(man_names, drop)


pattern <- "^doc_[0-9]+?$"

rm(list = ls(pattern = pattern))
out <- list()
for (man_name in man_names) {
  log_info("watching {man_name}")
  out[[man_name]] <- process_manual_office(name = man_name, pkg = "officer")
}

out <- Filter(function(x) length(x)>0, out)

stack_images(out$body_add$doc_1, fileout = "inst/medias/figures/body_add_doc_1.png")
stack_images(out$ph_with$doc_1, fileout = "inst/medias/figures/ph_with_doc_1.png")

# pdftools::pdf_convert(pdf = "~/Desktop/docx_prop_section_1_.pdf", format = "png")
# word_as_png(c("docx_prop_section_1__1.png", "docx_prop_section_1__2.png"), file = "inst/medias/figures/prop_section_doc_1.png")
# zz=pdftools::pdf_convert(pdf = "~/Desktop/body_end_block_section_doc_1.pdf", format = "png")
# word_as_png(zz, file = "inst/medias/figures/body_end_block_section_doc_1.png")
#
rsvg::rsvg_png("inst/medias/officerlogo.svg", "man/figures/logo.png")
img_compress(dir_input = "inst/medias/figures",
                       dir_output = "man/figures")


