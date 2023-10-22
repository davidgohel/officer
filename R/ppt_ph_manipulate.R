get_shape_id <- function(x, type = NULL, id = NULL, ph_label = NULL ){
  slsmry <- slide_summary(x)

  if( is.null(ph_label) ){
    sel <- which( slsmry$type %in% type )
    if( length(sel) < 1 ) stop("no shape of type ", shQuote(type), " has been found")
    sel <- sel[id]
    if( sum(is.finite(sel)) != 1 ) stop("no shape of type ", shQuote(type), " and with id ", id, " has been found")
  } else {
    sel <- which( slsmry$ph_label %in% ph_label )
    if( length(sel) < 1 ) stop("no shape with label ", shQuote(ph_label), " has been found")
    sel <- sel[id]
    if( sum(is.finite(sel)) != 1 ) stop("no shape with label ", shQuote(ph_label), "and with id ", id, " has been found")
  }
  sel
}




#' @export
#' @title Remove a shape
#' @description Remove a shape in a slide.
#' @param x an rpptx object
#' @param type placeholder type
#' @param id placeholder index (integer) for a duplicated type. This is to be used when a placeholder
#' type is not unique in the layout of the current slide, e.g. two placeholders with type 'body'. To
#' add onto the first, use \code{id = 1} and \code{id = 2} for the second one.
#' Values can be read from \code{\link{slide_summary}}.
#' @param ph_label label associated to the placeholder. Use column
#' \code{ph_label} of result returned by \code{\link{slide_summary}}.
#' @param id_chr deprecated.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' dummy_fun <- function(doc){
#'   doc <- add_slide(doc, layout = "Two Content",
#'     master = "Office Theme")
#'   doc <- ph_with(x = doc, value = "Un titre",
#'     location = ph_location_type(type = "title"))
#'   doc <- ph_with(x = doc, value = "Un corps 1",
#'     location = ph_location_type(type = "body", id = 1))
#'   doc <- ph_with(x = doc, value = "Un corps 2",
#'     location = ph_location_type(type = "body", id = 2))
#'   doc
#' }
#' doc <- read_pptx()
#' for(i in 1:3)
#'   doc <- dummy_fun(doc)
#'
#' doc <- on_slide(doc, index = 1)
#' doc <- ph_remove(x = doc, type = "title")
#'
#' doc <- on_slide(doc, index = 2)
#' doc <- ph_remove(x = doc, type = "body", id = 2)
#'
#' doc <- on_slide(doc, index = 3)
#' doc <- ph_remove(x = doc, type = "body", id = 1)
#'
#' print(doc, target = fileout )
#' @family functions for placeholders manipulation
#' @seealso \code{\link{ph_with}}
ph_remove <- function( x, type = "body", id = 1, ph_label = NULL, id_chr = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr][%.0f]", office_id) )

  xml_remove(current_elt)

  x
}



#' @export
#' @title Slide link to a placeholder
#' @description Add slide link to a placeholder in the current slide.
#' @inheritParams ph_remove
#' @param slide_index slide index to reach
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' loc_title <- ph_location_type(type = "title")
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, "Un titre 1", location = loc_title)
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, "Un titre 2", location = loc_title)
#' doc <- on_slide(doc, 1)
#' slide_summary(doc) # read column ph_label here
#' doc <- ph_slidelink(x = doc, ph_label = "Title 1", slide_index = 2)
#'
#' print(doc, target = fileout )
#' @family functions for placeholders manipulation
#' @seealso \code{\link{ph_with}}
ph_slidelink <- function( x, type = "body", id = 1, id_chr = NULL, ph_label = NULL, slide_index){

  slide <- x$slide$get_slide(x$cursor)
  office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr][%.0f]", office_id) )

  # declare slide ref in relationships
  slide_name <- x$slide$names()[slide_index]
  slide$reference_slide(slide_name)
  rel_df <- slide$rel_df()
  id <- rel_df[rel_df$target == slide_name, "id" ]

  # add hlinkClick
  cnvpr <- xml_child(current_elt, "p:nvSpPr/p:cNvPr")
  str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\" action=\"ppaction://hlinksldjump\"/>"
  str_ <- sprintf(str_, id)
  xml_add_child(cnvpr, as_xml_document(str_) )

  x
}


#' @export
#' @title Hyperlink a placeholder
#' @description Add hyperlink to a placeholder in the current slide.
#' @inheritParams ph_remove
#' @param href hyperlink (do not forget http or https prefix)
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' loc_manual <- ph_location(bg = "red", newlabel= "mytitle")
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(x = doc, "Un titre 1", location = loc_manual)
#' slide_summary(doc) # read column ph_label here
#' doc <- ph_hyperlink(x = doc, ph_label = "mytitle",
#'   href = "https://cran.r-project.org")
#'
#' print(doc, target = fileout )
#' @family functions for placeholders manipulation
#' @seealso \code{\link{ph_with}}
ph_hyperlink <- function( x, type = "body", id = 1, id_chr = NULL, ph_label = NULL, href ){

  slide <- x$slide$get_slide(x$cursor)
  office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr][%.0f]", office_id) )

  # add hlinkClick
  cnvpr <- xml_child(current_elt, "p:nvSpPr/p:cNvPr")
  str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\"/>"
  str_ <- sprintf(str_, href)
  xml_add_child(cnvpr, as_xml_document(str_) )
  x
}

