unlink("vignettes/assets", recursive = TRUE, force = TRUE)
pkgdown::build_site()
browseURL("docs/articles/assets/docx/toc_and_captions.docx")
