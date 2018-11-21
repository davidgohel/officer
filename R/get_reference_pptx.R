
get_default_pandoc_pptx <- function(outfile = tempfile(fileext = ".pptx")) {

  system2(rmarkdown::pandoc_exec(),
          args = c("--print-default-data-file", "reference.pptx"),
          stdout = outfile
  )
  return(outfile)
}

#' @export
#' @title Get the document being used as a template
#' @description Get filename of the PowerPoint document
#' being used as a template in an R Markdown document
#' render as PowerPoint presentation. It requires
#' packages rmarkdown >= 1.10.14 and knitr.
get_reference_pptx <- function() {

  if( !requireNamespace("rmarkdown") )
    stop("package rmarkdown >= 1.10.14 is required to use this function")
  if( !requireNamespace("knitr") )
    stop("package knitr is required to use this function")

  if( compareVersion(as.character(packageVersion("rmarkdown")), "1.10.14") < 0 )
    stop("package rmarkdown >= 1.10.14 is required to use this function")

  # if( compareVersion(as.character(rmarkdown::pandoc_version()), "2.4") < 0 )
  #   stop("package pandoc >= 2.4 is required to use this function")

  pandoc_args <- knitr::opts_knit$get("rmarkdown.pandoc.args")

  rd <- grep("--reference-doc", pandoc_args)
  if (length(rd)) {
    reference_pptx <- pandoc_args[rd + 1]
  } else {
    reference_pptx <- get_default_pandoc_pptx()
  }
  return(reference_pptx)
}

