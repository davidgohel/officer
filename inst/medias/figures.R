library(imgpress)
library(officer)

man_names <- package_man_names(package_name = "officer")
man_names <- c("body_add", "ph_with", "body_set_default_section",
               "read_docx", "body_end_block_section", "prop_section")

pattern <- "^doc_[0-9]+?$"

rm(list = ls(pattern = pattern))
out <- list()
for (man_name in man_names) {
  out[[man_name]] <- process_manual_office(name = man_name, pkg = "officer")
}

out <- Filter(function(x) length(x)>0, out)

out[["body_end_block_section"]] <- process_manual_office(name = "body_end_block_section", pkg = "officer")
out[["prop_section"]] <- process_manual_office(name = "prop_section", pkg = "officer")


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
                    fileout = "inst/medias/figures/
                    "
                    )
images_to_miniature(out$read_docx$doc_1,
                    width = 650, row = c(1, 1),
                    fileout = "inst/medias/figures/read_docx_doc_1.png"
                    )
images_to_miniature(out$read_docx$doc_2,
                    width = 850,
                    fileout = "inst/medias/figures/read_docx_doc_2.png"
                    )
images_to_miniature(out$body_end_block_section$doc_1,
                    width = 850,
                    fileout = "inst/medias/figures/body_end_block_section_doc_1.png"
                    )
images_to_miniature(out$prop_section$doc_1,
                    width = 850,
                    fileout = "inst/medias/figures/prop_section_doc_1.png"
                    )

rsvg::rsvg_png("inst/medias/officerlogo.svg", "man/figures/logo.png")
img_compress(dir_input = "inst/medias/figures",
                       dir_output = "man")


