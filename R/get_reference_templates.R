get_default_pandoc_data_file <- function(format = "pptx", outfile = tempfile(fileext = ".pptx")) {
  ref_doc <- paste0("reference.", format)
  system2(rmarkdown::pandoc_exec(),
          args = c("--print-default-data-file", ref_doc),
          stdout = outfile)
  return(outfile)
}

#' @export
#' @title Get the document being used as a template
#' @description Get filename of the document
#' being used as a template in an R Markdown document
#' rendered as HTML, PowerPoint presentation or Word document. It requires
#' packages rmarkdown >= 1.10.14 and knitr.
#' @param format document format, one of 'pptx', 'docx' or 'html'
#' @return a name file
#' @author Noam Ross
#' @importFrom utils compareVersion packageVersion
get_reference_value <- function(format = NULL) {

  if( length(format) != 1 ){
    stop("format must be a scalar character")
  }


  if( !requireNamespace("rmarkdown") )
    stop("package rmarkdown is required to use this function")
  if( !requireNamespace("knitr") )
    stop("package knitr is required to use this function")

  if( compareVersion(as.character(packageVersion("rmarkdown")), "1.10.14") < 0 )
    stop("package rmarkdown >= 1.10.14 is required to use this function")

  if( is.null(knitr::opts_knit$get("rmarkdown.pandoc.to"))){
    stop("Use this function within R markdown document")
  }
  if( is.null(format)){
    if( grepl( "docx", knitr::opts_knit$get("rmarkdown.pandoc.to") ) ){
      format <- "docx"
    } else if( grepl( "pptx", knitr::opts_knit$get("rmarkdown.pandoc.to") ) ){
      format <- "pptx"
    } else if( grepl( "html", knitr::opts_knit$get("rmarkdown.pandoc.to") ) ){
      format <- "html"
    } else if( grepl( "latex", knitr::opts_knit$get("rmarkdown.pandoc.to") ) ){
      format <- "latex"
    } else {
      stop("Unable to determine the format that should be used")
    }


  }
  if( !format %in% c("pptx", "docx", "html") ){
    stop("format must be have value 'docx', 'pptx' or 'html'.")
  }

  pandoc_args <- knitr::opts_knit$get("rmarkdown.pandoc.args")

  rd <- grep("--reference-doc", pandoc_args)
  if (length(rd)) {
    reference_data <- pandoc_args[rd + 1]
  } else {
    reference_data <- get_default_pandoc_data_file(format = format)
  }
  return(reference_data)
}

