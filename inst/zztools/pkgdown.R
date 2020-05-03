# unlink("vignettes/offcran/assets", recursive = TRUE, force = TRUE)
pkgdown::build_site()
file.copy("vignettes/offcran/assets", to = "docs/articles/offcran", overwrite = TRUE, recursive = TRUE)

browseURL("docs/articles/offcran/assets/docx/toc_and_captions.docx")
browseURL("docs/articles/offcran/assets/docx/body_add_demo.docx")

unlink("vignettes/offcran/assets", recursive = TRUE, force = TRUE)
unlink("vignettes/offcran/extract.png", recursive = TRUE, force = TRUE)
