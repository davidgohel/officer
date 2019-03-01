#' @export
#' @title append text
#' @description append text in a placeholder
#' @inheritParams ph_remove
#' @param str text to add
#' @param style text style, a \code{\link{fp_text}} object
#' @param pos where to add the new element relative to the cursor,
#' "after" or "before".
#' @param href hyperlink to reach when clicking the text
#' @param slide_index slide index to reach when clicking the text.
#' It will be ignored if \code{href} is not NULL.
#' @examples
#' library(magrittr)
#' fileout <- tempfile(fileext = ".pptx")
#' my_pres <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_empty(type = "body")
#'
#' small_red <- fp_text(color = "red", font.size = 14)
#'
#' my_pres <- my_pres %>%
#'   ph_add_par(level = 3) %>%
#'   ph_add_text(str = "A small red text.", style = small_red) %>%
#'   ph_add_par(level = 2) %>%
#'   ph_add_text(str = "Level 2")
#'
#' print(my_pres, target = fileout)
#'
#'
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_text(type = "title", str = "Un titre 1") %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_text(type = "title", str = "Un titre 2") %>%
#'   on_slide(1) %>%
#'   ph_empty(type = "body") %>%
#'   ph_add_par(type = "body", level = 2) %>%
#'   ph_add_text(str = "Jump here to slide 2!", type = "body",
#'   slide_index = 2)
#'
#' print(doc, target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
ph_add_text <- function( x, str, type = NULL, id_chr = NULL, ph_label = NULL,
  style = fp_text(font.size = 0), pos = "after",
  href = NULL, slide_index = NULL ){

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

  current_p <- xml_child(current_elt, "/a:p[last()]")
  if( inherits(current_p, "xml_missing") )
    stop("Could not find any paragraph in the selected shape.")

  r_shape_ <- sprintf(paste0( pml_with_ns("a:r"), "%s<a:t>%s</a:t></a:r>" ),
          format(style, type = "pml"),
          htmlEscape(str))

  if( pos == "after" )
    where_ <- length(xml_children(current_p))
  else where_ <- 0

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
#' @description append a new empty paragraph in a placeholder
#' @inheritParams ph_remove
#' @param level paragraph level
#' @examples
#' library(magrittr)
#'
#' fileout <- tempfile(fileext = ".pptx")
#' default_text <- fp_text(font.size = 0, bold = TRUE, color = "red")
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_text(type = "body", str = "A text") %>%
#'   ph_add_par(level = 2) %>%
#'   ph_add_text(str = "and another, ", style = default_text ) %>%
#'   ph_add_par(level = 3) %>%
#'   ph_add_text(str = "and another!",
#'     style = update(default_text, color = "blue"))
#'
#' print(doc, target = fileout)
#' @importFrom xml2 xml_child xml_children xml_add_child
ph_add_par <- function( x, type = NULL, id_chr = NULL, level = 1, ph_label = NULL ){

  slide <- x$slide$get_slide(x$cursor)

  if( is.null(ph_label) ){
    shape_id <- get_shape_id(x, type = type, id_chr = id_chr )
    nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/"))
    current_elt <- nodes[[shape_id]]
  } else {
    slsmry <- slide_summary(x)
    id_chr <- slsmry$id[slsmry$ph_label %in% ph_label]

    current_elt <- xml_find_first(slide$get(), sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", id_chr) )
  }
  current_p <- xml_child(current_elt, "/p:txBody")

  if( inherits(current_p, "xml_missing") ){
    if( level > 1 )
      p_shape <- sprintf("<a:p><a:pPr lvl=\"%.0f\"/></a:p>", level - 1)
    else
      p_shape <- "<a:p/>"

    simple_shape <- paste0( pml_with_ns("p:txBody"), "<a:bodyPr/><a:lstStyle/>",
                            p_shape, "</p:txBody>")
    xml_add_child(current_elt, as_xml_document(simple_shape) )
  } else {
    if( level > 1 ){
      simple_shape <- sprintf(paste0( pml_with_ns("a:p"), "<a:pPr lvl=\"%.0f\"/></a:p>" ), level - 1)
    } else
      simple_shape <- paste0(pml_with_ns("a:p"), "</a:p>")
    xml_add_child(current_p, as_xml_document(simple_shape) )
  }
  x
}



#' @export
#' @title append fpar
#' @description append \code{fpar} (a formatted paragraph) in a placeholder
#' @inheritParams ph_remove
#' @param value fpar object
#' @param level paragraph level
#' @param par_default specify if the default paragraph formatting
#' should be used.
#' @examples
#' library(magrittr)
#'
#' bold_face <- shortcuts$fp_bold(font.size = 30)
#' bold_redface <- update(bold_face, color = "red")
#'
#' fpar_ <- fpar(ftext("Hello ", prop = bold_face),
#'   ftext("World", prop = bold_redface ),
#'   ftext(", how are you?", prop = bold_face ) )
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_empty(type = "body") %>%
#'   ph_add_fpar(value = fpar_, type = "body", level = 2)
#'
#' print(doc, target = tempfile(fileext = ".pptx"))
#' @importFrom xml2 xml_child xml_children xml_add_child
#' @seealso \code{\link{fpar}}
ph_add_fpar <- function( x, value, type = "body", id_chr = NULL, ph_label = NULL,
                         level = 1, par_default = TRUE ){

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

  current_p <- xml_child(current_elt, "/p:txBody")
  newp_str <- format(value, type = "pml")
  newp_str <- gsub("<a:p>", pml_with_ns("a:p"), newp_str )
  node <- as_xml_document(newp_str)
  if( par_default ){
    ppr <- xml_child(node, "/a:pPr")
    empty_par <- as_xml_document(paste0(pml_with_ns("a:pPr"), "</a:pPr>"))
    xml_replace(ppr, empty_par )
  }
  ppr <- xml_child(node, "/a:pPr")
  if( level > 1 ){
    xml_attr(ppr, "lvl") <- sprintf("%.0f", level - 1)
  }
  if( inherits(current_p, "xml_missing") ){
    simple_shape <- paste0( pml_with_ns("p:txBody"), "<a:bodyPr/><a:lstStyle/></p:txBody>")
    newnode <- as_xml_document(simple_shape)
    xml_add_child(newnode, node)
    xml_add_child(current_elt, newnode )
  } else {
    xml_add_child(current_p, node )
  }

  x
}




