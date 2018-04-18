library(purrr)
library(stringr)

unlink("vignettes/assets", recursive = TRUE, force = TRUE)
pkgdown::build_site()

browseURL("docs/articles/assets/docx/toc_and_captions.docx")
browseURL("docs/articles/assets/docx/body_add_demo.docx")
browseURL("docs/articles/assets/docx/slip_in_demo.docx")

all_files <- list.files(path = "docs/reference", pattern = "\\.html", full.names = TRUE)
walk(all_files, .f = function(file){
  content <- readLines(file, encoding = "UTF-8")
  content <- str_replace_all(content, "/Users/davidgohel/Github/officer/docs/reference/", "/.../")
  cat(content, file = file, sep = "\n")
})

all_files <- list.files(path = "docs/articles", pattern = "\\.html", full.names = TRUE)
walk(all_files, .f = function(file){
  content <- readLines(file, encoding = "UTF-8")
  content <- str_replace_all(content, "/Users/davidgohel/Github/officer/docs/articles/", "/.../")
  cat(content, file = file, sep = "\n")
})

all_files <- list.files(path = "docs/articles/offcran", pattern = "\\.html", full.names = TRUE)
walk(all_files, .f = function(file){
  content <- readLines(file, encoding = "UTF-8")
  content <- str_replace_all(content, "/Users/davidgohel/Github/officer/docs/articles/assets/", "/.../")
  cat(content, file = file, sep = "\n")
})
