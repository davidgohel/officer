#' @export
#' @title open a connexion to a 'PowerPoint' file
#' @description read and import a pptx file as an R object
#' representing the document.
#' @param path path to the pptx file to use a base document.
#' @param x an rpptx object
#' @examples
#' read_pptx()
#' @importFrom xml2 read_xml xml_length xml_find_first xml_attr xml_ns
read_pptx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.pptx")

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- structure(list(package_dir = package_dir),
                   .Names = c("package_dir"),
                   class = "rpptx")

  obj$table_styles <- read_table_style(package_dir)

  obj$presentation <- presentation$new(obj)
  obj$slideLayouts <- dir_layout$new(obj )

  slide_xfrm <- obj$slideLayouts$xfrm()
  master_xfrm <- obj$slideLayouts$get_master()$xfrm()

  obj$slide <- dir_slide$new(obj )
  obj$content_type <- content_type$new(obj )

  obj$core_properties <- core_properties$new(obj$package_dir)

  obj$cursor = obj$slide$length()
  obj
}

read_table_style <- function(path){
  file <- file.path(path, "ppt/tableStyles.xml")
  doc <- read_xml(file)
  nodes <- xml_find_all(doc, "//a:tblStyleLst")
  data.frame(def = nodes %>% xml_attr("def"),
             styleName = nodes %>% xml_attr("styleName"),
             stringsAsFactors = FALSE )
}

#' @export
#' @param target path to the pptx file to write
#' @param ... unused
#' @rdname read_pptx
#' @examples
#' # write a rdocx object in a docx file ----
#' if( require(magrittr) ){
#'   read_pptx() %>% print(target = "out.pptx")
#'   # full path of produced file is returned
#'   print(.Last.value)
#' }
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

  x$core_properties$set_last_modified(format( Sys.time(), "%Y-%m-%dT%H:%M:%SZ"))
  x$core_properties$set_modified_by(Sys.getenv("USER"))
  x$core_properties$save()

  pack_folder(folder = x$package_dir, target = target )
}




#' @export
#' @importFrom lazyeval interp
#' @importFrom dplyr filter_
#' @importFrom xml2 xml_name<- xml_set_attrs xml_ns xml_remove
#' @title add a slide
#' @description add a slide into a pptx presentation
#' @param x rpptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres,
#'   layout = "Two Content", master = "Office Theme")
add_slide <- function( x, layout, master ){

  filter_criteria <- interp(~ name == layout & master_name == master, layout = layout, master = master)
  slide_info <- filter_(x$slideLayouts$get_metadata(), filter_criteria)
  if( nrow( slide_info ) < 1 )
    stop("could not find layout named ", shQuote(layout), " in master named ", shQuote(master))
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
#' @rdname read_pptx
#' @section number of slides:
#' Function \code{length} will return the number of slides.
length.rpptx <- function( x ){
  x$slide$length()
}

#' @export
#' @title change current slide
#' @description change current slide index of an rpptx object.
#' @param x rpptx object
#' @param index slide index
#' @examples
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- on_slide( doc, index = 1)
#' doc <- ph_with_text(x = doc, type = "title", str = "First title")
#' doc <- on_slide( doc, index = 3)
#' doc <- ph_with_text(x = doc, type = "title", str = "Third title")
#'
#' print(doc, target = "on_slide.pptx" )
on_slide <- function( x, index ){

  l_ <- length(x)
  if( l_ < 1 ){
    stop("presentation contains no slide", call. = FALSE)
  }
  if( !between(index, 1, l_ ) ){
    stop("unvalid index ", index, " (", l_," slide(s))", call. = FALSE)
  }

  x$cursor = index
  x
}



#' @export
#' @title remove a slide
#' @description remove a slide from a pptx presentation
#' @param x rpptx object
#' @param index slide index, default to current slide position.
#' @note cursor is set on the last slide.
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres,
#'   layout = "Two Content", master = "Office Theme")
#'
#' my_pres <- remove_slide(my_pres)
remove_slide <- function( x, index = NULL ){

  l_ <- length(x)
  if( l_ < 1 ){
    stop("presentation contains no slide to delete", call. = FALSE)
  }

  if( is.null(index) )
    index <- x$cursor

  if( !between(index, 1, l_ ) ){
    stop("unvalid index ", index, " (", l_," slide(s))", call. = FALSE)
  }

  del_file <- x$slide$remove(index)
  # update presentation elements
  x$presentation$remove_slide(del_file)
  x$content_type$remove_slide(partname = del_file )
  x$cursor = x$slide$length()
  x

}

#' @export
#' @title presentation layouts summary
#' @description get informations about slide layouts and
#' master layouts into a data.frame.
#' @param x rpptx object
#' @examples
#' my_pres <- read_pptx()
#' layout_summary ( x = my_pres )
layout_summary <- function( x ){
  data <- x$slideLayouts$get_metadata() %>%
    rename_( .dots = setNames( c("name", "master_name"), c("layout", "master"))) %>%
    select_("layout", "master")
  data
}

#' @export
#' @title slide layout properties
#' @description get informations about a particular slide layout
#' into a data.frame.
#' @param x rpptx object
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @examples
#' x <- read_pptx()
#' layout_properties ( x = x, layout = "Title Slide", master = "Office Theme" )
#' layout_properties ( x = x, master = "Office Theme" )
#' layout_properties ( x = x, layout = "Two Content" )
#' layout_properties ( x = x )
layout_properties <- function( x, layout = NULL, master = NULL ){

  data <- x$slideLayouts$get_xfrm_data()

  if( !is.null(layout) && !is.null(master) ){
    filter_criteria <- interp(~ name == layout & master_name == master, layout = layout, master = master)
    data <- filter_(data, filter_criteria)
  } else if( is.null(layout) && !is.null(master) ){
    filter_criteria <- interp(~ master_name == master, master = master)
    data <- filter_(data, filter_criteria)
  } else if( !is.null(layout) && is.null(master) ){
    filter_criteria <- interp(~ name == layout, layout = layout)
    data <- filter_(data, filter_criteria)
  }

  data <- select_(data, "master_name", "name", "type", "offx", "offy", "cx", "cy")

  data
}


#' @export
#' @title get PowerPoint slide content in a tidy format
#' @description get content and positions of current slide
#' into a data.frame. If any table, image or paragraph, data is
#' imported into the resulting data.frame.
#' @param x rpptx object
#' @param index slide index
#' @examples
#' library(magrittr)
#'
#' my_pres <- read_pptx() %>%
#'   add_slide(layout = "Two Content", master = "Office Theme") %>%
#'   ph_with_text(type = "dt", str = format(Sys.Date())) %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme")
#'
#' slide_summary(my_pres)
#' slide_summary(my_pres, index = 1)
#' @importFrom purrr map2_df
slide_summary <- function( x, index = NULL ){

  l_ <- length(x)
  if( l_ < 1 ){
    stop("presentation contains no slide", call. = FALSE)
  }

  if( is.null(index) )
    index <- x$cursor

  if( !between(index, 1, l_ ) ){
    stop("unvalid index ", index, " (", l_," slide(s))", call. = FALSE)
  }

  slide <- x$slide$get_slide(index)
  str = "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"
  nodes <- xml_find_all(slide$get(), str)
  data <- read_xfrm(nodes, file = "slide", name = "" )
  data$text <- map_chr(nodes, xml_text )
  select_(data, "-name", "-file", "-ph")
}






#' @export
#' @title color scheme
#' @description get master layout color scheme into a data.frame.
#' @param x rpptx object
#' @examples
#' x <- read_pptx()
#' color_scheme ( x = x )
color_scheme <- function( x ){
  x$slideLayouts$get_master()$get_color_scheme()
}
