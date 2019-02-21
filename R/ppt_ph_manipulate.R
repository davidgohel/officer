#' @export
#' @title remove shape
#' @description remove a shape in a slide
#' @param x a pptx device
#' @param type placeholder type
#' @param id_chr placeholder id (a string). This is to be used when a placeholder type
#' is not unique in the current slide, e.g. two placeholders with type 'body'.
#' Values can be read from \code{\link{slide_summary}}.
#' @param ph_label label associated to the placeholder. Use column
#' \code{ph_label} of result returned by \code{\link{slide_summary}}.
#' This should be used instead of using \code{type} and \code{id_chr}.
#' @note
#' Arguments \code{type} and \code{id_chr} will be deprecated in the next
#' version of flextable.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre")
#' slide_summary(doc) # read column ph_label here
#' doc <- ph_remove(x = doc, ph_label = "Title 1")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_remove xml_find_all
ph_remove <- function( x, type = NULL, id_chr = NULL, ph_label = NULL ){

  slide <- x$slide$get_slide(x$cursor)

  if( is.null(ph_label) ){
    shape_id <- get_shape_id(x, type = type, id_chr = id_chr )
    nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/"))
    current_elt <- nodes[[shape_id]]
  } else {
    slsmry <- slide_summary(x)
    id_chr <- slsmry$id[slsmry$ph_label %in% ph_label]# care if multiple ""
    current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", id_chr) )
  }

  xml_remove(current_elt)

  x
}



#' @export
#' @title slide link to a placeholder
#' @description add slide link to a placeholder in the current slide.
#' @inheritParams ph_remove
#' @param slide_index slide index to reach
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre 1")
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre 2")
#' doc <- on_slide(doc, 1)
#' slide_summary(doc) # read column ph_label here
#' doc <- ph_slidelink(x = doc, ph_label = "Title 1", slide_index = 2)
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_remove xml_find_all
ph_slidelink <- function( x, type = NULL, id_chr = NULL, ph_label = NULL, slide_index){

  slide <- x$slide$get_slide(x$cursor)
  if( is.null(ph_label) ){
    shape_id <- get_shape_id(x, type = type, id_chr = id_chr )
    nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/"))
    current_elt <- nodes[[shape_id]]
  } else {
    slsmry <- slide_summary(x)
    id_chr <- slsmry$id[slsmry$ph_label %in% ph_label]# care if multiple ""
    current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", id_chr) )
  }

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
#' @title hyperlink a placeholder
#' @description add hyperlink to a placeholder in the current slide.
#' @inheritParams ph_remove
#' @param href hyperlink (do not forget http or https prefix)
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with_text(x = doc, type = "title", str = "Un titre 1")
#' slide_summary(doc) # read column ph_label here
#' doc <- ph_hyperlink(x = doc, ph_label = "Title 1",
#'   href = "https://cran.r-project.org")
#'
#' print(doc, target = fileout )
#' @importFrom xml2 xml_remove xml_find_all
ph_hyperlink <- function( x, type = NULL, id_chr = NULL, ph_label = NULL, href ){

  slide <- x$slide$get_slide(x$cursor)
  if( is.null(ph_label) ){
    shape_id <- get_shape_id(x, type = type, id_chr = id_chr )
    nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/"))
    node <- nodes[[shape_id]]
  } else {
    slsmry <- slide_summary(x)
    id_chr <- slsmry$id[slsmry$ph_label %in% ph_label]# care if multiple ""
    node <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", id_chr) )
  }

  # declare link in relationships
  slide$reference_hyperlink(href)
  rel_df <- slide$rel_df()
  id <- rel_df[rel_df$target == href, "id" ]

  # add hlinkClick
  cnvpr <- xml_child(node, "p:nvSpPr/p:cNvPr")
  str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\"/>"
  str_ <- sprintf(str_, id)
  xml_add_child(cnvpr, as_xml_document(str_) )
  x
}

