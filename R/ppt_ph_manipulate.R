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
  slsmry$id[sel]
}




#' @export
#' @title remove a shape
#' @description remove a shape in a slide
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
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[*/p:cNvPr[@id='%s']]", office_id) )

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
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

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
  node <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

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

#' @export
#' @title append text
#' @description append text in a placeholder.
#' The function let you add text to an existing
#' content in an exiisting shape, existing text will be preserved.
#' @section Usage:
#' If your goal is to add formatted text in a new shape, use \code{\link{ph_with}}
#' with a \code{\link{block_list}} instead of this function.
#' @inheritParams ph_remove
#' @param str text to add
#' @param style text style, a \code{\link{fp_text}} object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @param href hyperlink to reach when clicking the text
#' @param slide_index slide index to reach when clicking the text.
#' It will be ignored if \code{href} is not NULL.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres)
#' my_pres <- ph_with(my_pres, "",
#'   location = ph_location_type(type = "body"))
#'
#' small_red <- fp_text(color = "red", font.size = 14)
#'
#' my_pres <- ph_add_text(my_pres, str = "A small red text.",
#'   style = small_red)
#' my_pres <- ph_add_par(my_pres, level = 2)
#' my_pres <- ph_add_text(my_pres, str = "Level 2")
#'
#' print(my_pres, target = fileout)
#'
#' # another example ----
#' fileout <- tempfile(fileext = ".pptx")
#'
#' doc <- read_pptx()
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Un titre 2",
#'   location = ph_location_type(type = "title"))
#' doc <- ph_with(doc, "",
#'   location = ph_location(rotation = 90, bg = "red",
#'       newlabel = "myph"))
#' doc <- ph_add_text(doc, str = "dummy text",
#'   ph_label = "myph")
#'
#' print(doc, target = fileout)
ph_add_text <- function( x, str, type = "body", id = 1, id_chr = NULL, ph_label = NULL,
                         style = fp_text(font.size = 0), pos = "after",
                         href = NULL, slide_index = NULL ){

  slide <- x$slide$get_slide(x$cursor)
  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }

  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

  current_p <- xml_child(current_elt, "/a:p[last()]")
  if( inherits(current_p, "xml_missing") )
    stop("Could not find any paragraph in the selected shape.")

  r_shape_ <- to_pml(ftext(text = str, prop = style), add_ns = TRUE)

  if( pos == "after" ){
    runset <- xml_find_all(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]/a:p[last()]/a:r", office_id) )
    where_ <- length(runset)
  } else where_ <- 0

  new_node <- as_xml_document(r_shape_)

  if( !is.null(href)){
    slide$reference_hyperlink(href)
    rel_df <- slide$rel_df()
    id <- rel_df[rel_df$target == href, "id" ]

    apr <- xml_child(new_node, "a:rPr")
    str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\"/>"
    str_ <- sprintf(str_, id)
    xml_add_child(apr, as_xml_document(str_) )
  } else if( !is.null(slide_index)){
    slide_name <- x$slide$names()[slide_index]
    slide$reference_slide(slide_name)
    rel_df <- slide$rel_df()
    id <- rel_df[rel_df$target == slide_name, "id" ]
    # add hlinkClick
    apr <- xml_child(new_node, "a:rPr")
    str_ <- "<a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\" action=\"ppaction://hlinksldjump\"/>"
    str_ <- sprintf(str_, id)
    xml_add_child(apr, as_xml_document(str_) )
  }


  xml_add_child(current_p, new_node, .where = where_ )

  slide$fortify_id()$save()

  x
}

#' @export
#' @title append paragraph
#' @description append a new empty paragraph in a placeholder.
#' The function let you add a new empty paragraph to an existing
#' content in an exiisting shape, existing paragraphs will be preserved.
#' @inheritParams ph_remove
#' @param level paragraph level
#' @inheritSection ph_add_text Usage
#' @examples
#' library(magrittr)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' default_text <- fp_text(font.size = 0, bold = TRUE, color = "red")
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("A text", location = ph_location_type(type = "body")) %>%
#'   ph_add_par(level = 2) %>%
#'   ph_add_text(str = "and another, ", style = default_text ) %>%
#'   ph_add_par(level = 3) %>%
#'   ph_add_text(str = "and another!",
#'               style = update(default_text, color = "blue"))
#'
#' print(doc, target = fileout)
ph_add_par <- function( x, type = "body", id = 1, id_chr = NULL, level = 1, ph_label = NULL ){

  slide <- x$slide$get_slide(x$cursor)

  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )
  current_p <- xml_child(current_elt, "/p:txBody")

  if( inherits(current_p, "xml_missing") ){
    if( level > 1 )
      p_shape <- sprintf("<a:p><a:pPr lvl=\"%.0f\"/></a:p>", level - 1)
    else
      p_shape <- "<a:p/>"

    simple_shape <- paste0( "<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                            "<a:bodyPr/><a:lstStyle/>",
                            p_shape, "</p:txBody>")
    xml_add_child(current_elt, as_xml_document(simple_shape) )
  } else {
    if( level > 1 ){
      simple_shape <- sprintf(paste0( ap_ns_yes, "<a:pPr lvl=\"%.0f\"/></a:p>" ), level - 1)
    } else
      simple_shape <- paste0(ap_ns_yes, "</a:p>")
    xml_add_child(current_p, as_xml_document(simple_shape) )
  }
  x
}



#' @export
#' @title append fpar
#' @description append \code{fpar} (a formatted paragraph) in a placeholder
#' The function let you add a new formatted paragraph (\code{\link{fpar}})
#' to an existing content in an exiisting shape, existing paragraphs
#' will be preserved.
#' @inheritParams ph_remove
#' @param value fpar object
#' @param level paragraph level
#' @param par_default specify if the default paragraph formatting
#' should be used.
#' @inheritSection ph_add_text Usage
#' @examples
#' library(magrittr)
#'
#' bold_face <- shortcuts$fp_bold(font.size = 30)
#' bold_redface <- update(bold_face, color = "red")
#'
#' fpar_ <- fpar(ftext("Hello ", prop = bold_face),
#'               ftext("World", prop = bold_redface ),
#'               ftext(", how are you?", prop = bold_face ) )
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with("", location = ph_location(bg = "wheat", newlabel = "myph")) %>%
#'   ph_add_fpar(value = fpar_, ph_label = "myph", level = 2)
#'
#' print(doc, target = tempfile(fileext = ".pptx"))
#' @seealso \code{\link{fpar}}
ph_add_fpar <- function( x, value, type = "body", id = 1, id_chr = NULL, ph_label = NULL,
                         level = 1, par_default = TRUE ){

  slide <- x$slide$get_slide(x$cursor)

  if( !is.null(id_chr)) {
    office_id <- id_chr
  } else {
    office_id <- get_shape_id(x, type = type, id = id, ph_label = ph_label )
  }
  current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", office_id) )

  current_p <- xml_child(current_elt, "/p:txBody")

  node <- as_xml_document(to_pml(value, add_ns = TRUE))

  if( par_default ){
    # add default pPr
    ppr <- xml_child(node, "/a:pPr")
    empty_par <- as_xml_document(paste0("<a:pPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                                        "</a:pPr>"))
    xml_replace(ppr, empty_par )
  }
  ppr <- xml_child(node, "/a:pPr")
  if( level > 1 ){
    xml_attr(ppr, "lvl") <- sprintf("%.0f", level - 1)
  }
  if( inherits(current_p, "xml_missing") ){
    simple_shape <- paste0( "<p:txBody xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                            "<a:bodyPr/><a:lstStyle/></p:txBody>")
    newnode <- as_xml_document(simple_shape)
    xml_add_child(newnode, node)
    xml_add_child(current_elt, newnode )
  } else {
    xml_add_child(current_p, node )
  }

  x
}




