# unlink("vignettes/offcran/assets", recursive = TRUE, force = TRUE)
pkgdown::build_site()

browseURL("vignettes/offcran/assets/docx/toc_and_captions.docx")
browseURL("vignettes/offcran/assets/docx/body_add_demo.docx")
browseURL("vignettes/offcran/assets/docx/slip_in_demo.docx")

file.copy("vignettes/offcran/assets", to = "docs/articles/offcran", overwrite = TRUE, recursive = TRUE)
unlink("vignettes/offcran/assets", recursive = TRUE, force = TRUE)
unlink("vignettes/offcran/extract.png", recursive = TRUE, force = TRUE)
