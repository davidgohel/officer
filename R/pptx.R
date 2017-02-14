#' @export
#' @title open a connexion to a 'PowerPoint' file
#' @description read and import a pptx file as an R object
#' representing the document.
#' @param path path to the pptx file to use a base document.
#' @param x,object a pptx object
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

  slide_xfrm <- obj$slideLayouts$xfrm()
  master_xfrm <- obj$slideLayouts$get_master()$xfrm()
  obj$coordinates <- xfrmize( slide_xfrm, master_xfrm )

  obj$slide <- dir_slide$new(obj )
  obj$content_type <- content_type$new(obj )
  obj$cursor = obj$slide$length()
  obj
}



#' @export
#' @param target path to the pptx file to write
#' @param ... unused
#' @rdname pptx
#' @importFrom tools file_ext
print.pptx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("pptx document\n")
    print(x$coordinates[, c("type", "name", "master_file")])
    return(invisible())
  }

  if( file_ext(target) != "pptx" )
    stop(target , " should have '.pptx' extension.")

  x$presentation$save()
  x$content_type$save()
  pack_folder(folder = x$package_dir, target = target )
}




#' @export
#' @importFrom lazyeval interp
#' @importFrom dplyr filter_
#' @importFrom xml2 xml_name<- xml_set_attrs xml_ns xml_remove
#' @title add a slide
#' @description add a slide to a pptx presentation
#' @param x pptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
add_slide <- function( x, layout = "Titre et contenu", master = "masque1" ){

  filter_criteria <- interp(~ name == layout & master_name == master, layout = layout, master = master)
  slide_info <- filter_(x$slideLayouts$get_metadata(), filter_criteria)
  new_slidename <- x$slide$get_new_slidename()

  xml_file <- file.path(x$package_dir, "ppt/slides", new_slidename)
  xml_layout <- file.path(x$package_dir, "ppt/slideLayouts", slide_info$filename)

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

  # update presentation elements
  x$presentation$add_slide(target = file.path( "slides", new_slidename) )
  x$content_type$add_slide(partname = file.path( "/ppt/slides", new_slidename) )
  x$slide$update()
  x$cursor = x$slide$length()
  x

}

#' @export
#' @rdname pptx
#' @section number of slides:
#' Function \code{length} will return the number of slides.
length.pptx <- function( x ){
  x$slide$length()
}

#' @export
#' @title change current slide
#' @description change current slide index of a pptx object.
#' @param x pptx object
#' @param index slide index
#' @examples
#' doc <- pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- on_slide( doc, index = 1)
#' doc <- placeholder_set_text(x = doc, id = "title", str = "First title")
#' doc <- on_slide( doc, index = 3)
#' doc <- placeholder_set_text(x = doc, id = "title", str = "Third title")
#'
#' print(doc, target = "on_slide.pptx" )
on_slide <- function( x, index ){
  if(index > length(x) || index < 1)
    stop("invalid index value: ", index)
  x$cursor = index
  x
}


#' @export
#' @rdname pptx
#' @section slide summary:
#' Function \code{summary} will return a data.frame describing the content of
#' the current slide.
#' @importFrom xml2 xml_find_all xml_text
summary.pptx <- function( object, ... ){

  slide <- object$slide$get_slide(object$cursor)
  str = "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"
  nodes <- xml_find_all(slide$get(), str)
  xfrm <- read_xfrm(nodes, file = "slide", name = "" )
  xfrm$name <- NULL
  xfrm$file <- NULL
  xfrm$ph <- NULL
  xfrm$text <- xml_text(nodes)
  xfrm
}

#' @export
#' @title remove a slide
#' @description remove a slide from a pptx presentation
#' @param x pptx object
remove_slide <- function( x ){

  del_file <- x$slide$remove(x$cursor)
  # update presentation elements
  x$presentation$remove_slide(del_file)
  x$content_type$remove_slide(partname = del_file )
  x$cursor = x$slide$length()
  x

}


