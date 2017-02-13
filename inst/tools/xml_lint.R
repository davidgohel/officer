xml_lint <- function(fi) {

  file <- R.utils::filePath(fi)
  dest <- tempfile(fileext = ".xml")
  file.copy(from = file, to = dest)
  cmd <- paste("xmllint --format", dest, ">", file)
  system(cmd)
  fi
}

dir_xlint <- function( dir_ ){
  purrr::walk ( list.files(dir_, all.files = TRUE, pattern = "\\.(xml|xml\\.rels)$", recursive = TRUE, full.names = TRUE),
                xml_lint )
}
