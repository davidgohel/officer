library(imgpress)

man_names <- package_man_names(package_name = "officer")
man_names <- c("body_add", "ph_with", "body_set_default_section",
               "read_docx")

pattern <- "^doc_[0-9]+?$"

rm(list = ls(pattern = pattern))
out <- list()
for (man_name in man_names) {
  out[[man_name]] <- process_manual_office(name = man_name, pkg = "officer")
}

out <- Filter(function(x) length(x)>0, out)

images_to_miniature(out$body_add$doc_1,
                    row = c(1, 1, 2, 3),
                    width = 700,
                    fileout = "inst/medias/figures/body_add_doc_1.png")

images_to_miniature(out$ph_with$doc_1,
                    width = 800,
                    row = c(1, 1, 2, 2, 3, 3,4),
                    fileout = "inst/medias/figures/ph_with_doc_1.png"
                    )
images_to_miniature(out$body_set_default_section$doc_1,
                    width = 850,
                    fileout = "inst/medias/figures/body_set_default_section_doc_1.png"
                    )

# pdftools::pdf_convert(pdf = "~/Desktop/docx_prop_section_1_.pdf", format = "png")
# word_as_png(c("docx_prop_section_1__1.png", "docx_prop_section_1__2.png"), file = "inst/medias/figures/prop_section_doc_1.png")
# zz=pdftools::pdf_convert(pdf = "~/Desktop/body_end_block_section_doc_1.pdf", format = "png")
# word_as_png(zz, file = "inst/medias/figures/body_end_block_section_doc_1.png")
#
rsvg::rsvg_png("inst/medias/officerlogo.svg", "man/figures/logo.png")
img_compress(dir_input = "inst/medias/figures",
                       dir_output = "man")


