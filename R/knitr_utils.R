
check_dep <- function(){
  if( !requireNamespace("rmarkdown") )
    stop("package rmarkdown is required to use this function", call. = FALSE)
  if( !requireNamespace("knitr") )
    stop("package knitr is required to use this function", call. = FALSE)
}

get_default_pandoc_data_file <- function(format = "pptx") {
  outfile <- tempfile(fileext = paste0(".", format))

  pandoc_exec <- rmarkdown::pandoc_exec()
  if(.Platform$OS.type %in% "windows"){
    pandoc_exec <- paste0(pandoc_exec, ".exe")
  }

  if(!rmarkdown::pandoc_available() || !file.exists(pandoc_exec)){
    # Use officer template when no pandoc - stupid fallback as if o pandoc, no result
    outfile <- system.file(package = "officer", "template", paste0("template.", format))
  } else {
    # note in that case outfile will be removed when session will end
    ref_doc <- paste0("reference.", format)
    system2(pandoc_exec,
            args = c("--print-default-data-file", ref_doc),
            stdout = outfile)
  }

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
#' @importFrom utils compareVersion packageVersion
#' @family functions for officer extensions
#' @keywords internal
get_reference_value <- function(format = NULL) {

  if( !is.null(format) && length(format) != 1 ){
    stop("format must be a scalar character")
  }

  check_dep()

  if( compareVersion(as.character(packageVersion("rmarkdown")), "1.10.14") < 0 )
    stop("package rmarkdown >= 1.10.14 is required to use this function")

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

  output.dir <- knitr::opts_knit$get("output.dir")
  if(is.null(output.dir)){
    output.dir <- getwd()
  }

  pandoc_args <- knitr::opts_knit$get("rmarkdown.pandoc.args")

  rd <- grep("--reference-doc", pandoc_args)
  if (length(rd)) {
    reference_data <- pandoc_args[rd + 1]
    if(!file.exists(reference_data))
      reference_data <- file.path(output.dir, reference_data)
  } else {
    reference_data <- get_default_pandoc_data_file(format = format)
  }
  return(normalizePath(reference_data, winslash = "/"))
}


knitr_opts_current <- function(x, default = FALSE){
  check_dep()
  x <- knitr::opts_current$get(x)
  if(is.null(x)) x <- default
  x
}


#' @export
#' @title Table options in a 'knitr' context
#' @description Get options for table rendering in a 'knitr'
#' context. It should not be used by the end user, but
#' its documentation should be read as a place where
#' table options are documented when 'knitr' is used.
#'
#' The function is a utility to facilitate the retrieval of table
#' options supported by the 'flextable', 'officedown' and of
#' course 'officer' packages.
#'
#' These options should be set with `knitr::opts_chunk$set()`.
#' The names and expected values are listed in the following sections.
#'
#' @section knitr chunk options for table captions:
#'
#' | **label**                                        |    **name**     | **value**  |
#' |:-------------------------------------------------|:---------------:|:----------:|
#' | caption id/bookmark                              | tab.id          |    NULL    |
#' | caption                                          | tab.cap         |    NULL    |
#' | display table caption on top of the table or not | tab.topcaption  |    TRUE    |
#' | caption table sequence identifier.               | tab.lp          |   "tab:"   |
#'
#' @section knitr chunk options for Word table captions:
#'
#' | **label**                                               |    **name**     | **value**                 |
#' |:--------------------------------------------------------|:---------------:|:-------------------------:|
#' | Word stylename to use for table captions.               | tab.cap.style   |            NULL           |
#' | prefix for numbering chunk (default to   "Table ").     | tab.cap.pre     |           Table           |
#' | suffix for numbering chunk (default to   ": ").         | tab.cap.sep     |            " :"           |
#' | title number depth                                      | tab.cap.tnd     |              0            |
#' | separator to use between title number and table number. | tab.cap.tns     |             "-"           |
#' | caption prefix formatting properties                    | tab.cap.fp_text | fp_text_lite(bold = TRUE) |
#'
#' @section knitr chunk options for Word tables:
#'
#' | **label**                                                     |    **name**  | **value**  |
#' |:--------------------------------------------------------------|:------------:|:----------:|
#' | the Word stylename to use for tables                          | tab.style    |    NULL    |
#' | autofit' or 'fixed' algorithm.                                | tab.layout   |  "autofit" |
#' | value of the preferred width of the table in percent (base 1).| tab.width    |      1     |
#'
#' @section knitr chunk options for data.frame with officedown:
#'
#' | **label**                                                     |    **name**     | **value**  |
#' |:--------------------------------------------------------------|:---------------:|:----------:|
#' | apply or remove formatting from the first row in the table    | first_row       |    TRUE    |
#' | apply or remove formatting from the first column in the table | first_column    |    FALSE   |
#' | apply or remove formatting from the last row in the table     | last_row        |    FALSE   |
#' | apply or remove formatting from the last column in the table  | last_column     |    FALSE   |
#' | don't display odd and even rows                               | no_hband        |    TRUE    |
#' | don't display odd and even columns                            | no_vband        |    TRUE    |
#'
#' @return a list with following elements:
#'
#' * cap.style (default: NULL)
#' * cap.pre (default: "Table ")
#' * cap.sep (default: ":")
#' * cap.tnd (default: 0)
#' * cap.tns (default: "-")
#' * cap.fp_text (default: `fp_text_lite(bold = TRUE)`)
#' * id (default: NULL)
#' * cap (default: NULL)
#' * topcaption (default: TRUE)
#' * style (default: NULL)
#' * tab.lp (default: "tab:")
#' * table_layout (default: "autofit")
#' * table_width (default: 1)
#' * first_row (default: TRUE)
#' * first_column (default: FALSE)
#' * last_row (default: FALSE)
#' * last_column (default: FALSE)
#' * no_hband (default: TRUE)
#' * no_vband (default: TRUE)
#' @family functions for officer extensions
#' @keywords internal
opts_current_table <- function() {
  tab.cap.style <- knitr_opts_current("tab.cap.style", default = NULL)
  tab.cap.pre <- knitr_opts_current("tab.cap.pre", default = "Table ")
  tab.cap.sep <- knitr_opts_current("tab.cap.sep", default = ": ")

  tab.cap.tnd <- knitr_opts_current("tab.cap.tnd", default = 0)
  tab.cap.tns <- knitr_opts_current("tab.cap.tns", default = "-")
  tab.cap.fp_text <- knitr_opts_current("tab.cap.fp_text", default = fp_text_lite(bold = TRUE))

  tab.cap <- knitr_opts_current("tab.cap", default = NULL)
  tab.topcaption <- knitr_opts_current("tab.topcaption", default = TRUE)
  tab.id <- knitr_opts_current("tab.id", default = NULL)
  tab.lp <- knitr_opts_current("tab.lp", default = "tab:")
  tab.style <- knitr_opts_current("tab.style", default = NULL)
  tab.layout <- knitr_opts_current("tab.layout", default = "autofit")
  tab.width <- knitr_opts_current("tab.width", default = 1)

  first_row <- knitr_opts_current("first_row", default = TRUE)
  first_column <- knitr_opts_current("first_column", default = FALSE)
  last_row <- knitr_opts_current("last_row", default = FALSE)
  last_column <- knitr_opts_current("last_column", default = FALSE)
  no_hband <- knitr_opts_current("no_hband", default = TRUE)
  no_vband <- knitr_opts_current("no_vband", default = TRUE)

  list(
    cap.style = tab.cap.style,
    cap.pre = tab.cap.pre,
    cap.sep = tab.cap.sep,
    cap.tnd = tab.cap.tnd,
    cap.tns = tab.cap.tns,
    cap.fp_text = tab.cap.fp_text,
    id = tab.id,
    topcaption = tab.topcaption,
    cap = tab.cap,
    style = tab.style,
    tab.lp = tab.lp,
    table_layout = table_layout(type = tab.layout),
    table_width = table_width(width = tab.width, unit = "pct"),
    first_row = first_row,
    first_column = first_column,
    last_row = last_row,
    last_column = last_column,
    no_hband = no_hband,
    no_vband = no_vband
  )
}
