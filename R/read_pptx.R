#' @export
#' @title open a connexion to a 'PowerPoint' file
#' @description read and import a pptx file as an R object
#' representing the document.
#' The function is called `read_pptx` because it allows you to initialize an
#' object of class `rpptx` from an existing PowerPoint file. Content will be
#' added to the existing presentation. By default, an empty document is used.
#' @param path path to the pptx file to use as base document.
#' @section master layouts and slide layouts:
#' `read_pptx()` uses a PowerPoint file as the initial document. This is the original
#' PowerPoint document where all slide layouts, placeholders for shapes and
#' styles come from. Major points to be aware of are:
#'
#'   * Slide layouts are relative to a master layout. A document can contain one or
#' more master layouts; a master layout can contain one or more slide layouts.
#' * A slide layout inherits design properties from its master layout but some
#' properties can be overwritten.
#' * Designs and formatting properties of layouts and shapes (placeholders in a
#'                                                            layout) are defined within the initial document. There is no R function to
#' modify these values - they must be defined in the initial document.
#' @examples
#' read_pptx()
#' @importFrom xml2 read_xml xml_length xml_find_first xml_attr xml_ns
#' @seealso \code{\link{print.rpptx}} \code{\link{add_slide}}
read_pptx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.pptx")

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- list(package_dir = package_dir)


  obj$table_styles <- read_table_style(package_dir)

  obj$presentation <- presentation$new(package_dir)

  obj$masterLayouts <- dir_master$new(package_dir, slide_master$new("ppt/slideMasters") )

  obj$slideLayouts <- dir_layout$new( package_dir,
                                      master_metadata = obj$masterLayouts$get_metadata(),
                                      master_xfrm = obj$masterLayouts$xfrm() )

  obj$slide <- dir_slide$new( package_dir, obj$slideLayouts$get_xfrm_data() )
  obj$content_type <- content_type$new( package_dir )
  obj$core_properties <- read_core_properties(package_dir)

  obj$cursor = obj$slide$length()
  class(obj) <- "rpptx"
  obj
}

read_table_style <- function(path){
  file <- file.path(path, "ppt/tableStyles.xml")
  if (!file.exists(file)) {
    warning("tableStyles.xml file does not exist in PPTX")
    return(NULL)
  }
  doc <- read_xml(file)
  nodes <- xml_find_all(doc, "//a:tblStyleLst")
  data.frame(def = xml_attr(nodes, "def"),
             styleName = xml_attr(nodes, "styleName"),
             stringsAsFactors = FALSE )
}

#' write a 'PowerPoint' file.
#'
#'
#' @param x an rpptx object
#' @param target path to the pptx file to write
#' @param ... unused
#' @examples
#' # write a rdocx object in a docx file ----
#' file <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' print(doc, target = file)
#' @export
#' @seealso \code{\link{read_pptx}}
print.rpptx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("pptx document with", length(x), "slide(s)\n")
    cat("Available layouts and their associated master(s) are:\n")
    print(as.data.frame( layout_summary(x)) )
    return(invisible())
  }

  if( !grepl(x = target, pattern = "\\.(pptx)$", ignore.case = TRUE) )
    stop(target , " should have '.pptx' extension.")

  x$presentation$save()
  x$content_type$save()

  x$slide$save_slides()

  x$core_properties['modified','value'] <- format( Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
  x$core_properties['lastModifiedBy','value'] <- Sys.getenv("USER")
  write_core_properties(x$core_properties, x$package_dir)

  invisible(pack_folder(folder = x$package_dir, target = target ))
}

