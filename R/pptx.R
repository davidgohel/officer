#' @export
#' @title open a connexion to a 'PowerPoint' file
#' @description read and import a pptx file as an R object
#' representing the document.
#' @param path path to the pptx file to use a base document.
#' @param x a pptx object
#' @examples
#' # create a pptx object with default template ---
#' pptx()
#'
#' @importFrom xml2 read_xml xml_length xml_find_first xml_attr xml_ns
#' @importFrom utils unzip
pptx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.pptx")

  package_dir <- tempfile()
  unzip( zipfile = path, exdir = package_dir )

  obj <- structure(list(package_dir = package_dir),
                   .Names = c("package_dir"),
                   class = "pptx")
  obj$presentation <- presentation$new(obj)
  obj$slideLayouts <- dir_layout$new(obj )
  obj$slide <- dir_slide$new(obj )
  obj$content_type <- content_type$new(obj )
  obj$cursor = obj$slide$length()
  obj$coordinates <- xfrmize( obj )
  obj
}



#' @export
#' @param target path to the pptx file to write
#' @param ... unused
#' @rdname pptx
print.pptx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("pptx document\n")
    print(x$coordinates)
    return(invisible())
  }
  x$presentation$save()
  x$content_type$save()
  pack_folder(folder = x$package_dir, target = target )
}




#' @export
#' @importFrom lazyeval interp
#' @importFrom dplyr filter_
#' @importFrom xml2 xml_name<- xml_set_attrs xml_ns xml_remove
add_slide <- function( x, layout = "Titre et contenu", master = "masque1" ){

  filter_criteria <- interp(~ name == layout & master_name == master, layout = layout, master = master)
  slide_info <- filter_(x$slideLayouts$get_data(), filter_criteria)
  new_slidename <- x$slide$get_new_slidename()
  xml_file <- file.path(x$package_dir, "ppt/slides",
                        new_slidename)
  xml_layout <- file.path(x$package_dir, "ppt/slideLayouts",
                          slide_info$filename)

  xml_doc <- read_xml(xml_layout)
  xml_name(xml_doc) <- "sld"
  ns <- xml_ns(xml_doc)
  xml_attr(xml_doc, "type" ) <- NULL
  xml_attr(xml_doc, "preserve" ) <- NULL

  map(xml_find_all(xml_doc, "//p:sp"), xml_remove)


  write_xml(xml_doc, xml_file)

  rel_filename <- file.path(
    dirname(xml_file), "_rels",
    paste0(basename(xml_file), ".rels") )
  newrel <- relationship$new()$add(
    id = "rId1", type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    target = file.path("../slideLayouts", basename(slide_info$filename)) )
  newrel$write(path = rel_filename)

  # update preentation elements
  x$presentation$add_slide(target = file.path( "slides", new_slidename) )
  x$content_type$add_slide(partname = file.path( "/ppt/slides", new_slidename) )
  x$slide$update()
  x$cursor = x$slide$length()

  x

}


