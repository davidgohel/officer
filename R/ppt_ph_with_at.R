ph <- function( left = 0, top = 0, width = 3, height = 3, bg = "transparent", rot = 0){

  if( !is.color( bg ) )
    stop("bg must be a valid color.", call. = FALSE )
  cols <- as.integer( col2rgb(bg, alpha = TRUE)[,1] )

  p_ph(offx=as.integer( left * 914400 ),
    offy = as.integer( top * 914400 ),
    cx = as.integer( width * 914400 ),
    cy = as.integer( height * 914400 ),
    rot = as.integer(-rot * 60000),
    r = cols[1],
    g = cols[2],
    b = cols[3],
    a = cols[4] )
}

#' @rdname ph_empty
#' @export
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
#' @param bg background color
#' @param rot rotation angle
#' @examples
#'
#' # demo ph_empty_at ------
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_empty_at(x = doc, left = 1, top = 2, width = 5, height = 4)
#'
#' print(doc, target = fileout )
ph_empty_at <- function( x, left, top, width, height, bg = "transparent", rot = 0 ){

  slide <- x$slide$get_slide(x$cursor)

  new_ph <- ph(left = left, top = top, width = width, height = height, rot = rot, bg = bg)
  new_ph <- paste0( pml_with_ns("p:sp"), new_ph,"</p:sp>")
  new_node <- as_xml_document(new_ph)

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), new_node)

  slide$save()
  x$slide$update()
  x
}


#' @export
#' @rdname ph_with_img
#' @param left,top location of the new shape on the slide
#' @param rot rotation angle
#' @examples
#'
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#'
#' img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )
#' if( file.exists(img.file) ){
#'   doc <- ph_with_img_at(x = doc, src = img.file, height = 1.06, width = 1.39,
#'     left = 4, top = 4, rot = 45 )
#' }
#'
#' print(doc, target = fileout )
ph_with_img_at <- function( x, src, left, top, width, height, rot = 0 ){

  slide <- x$slide$get_slide(x$cursor)

  ext_img <- external_img(src, width = width, height = height)
  xml_elt <- format(ext_img, type = "pml")

  slide$reference_img(src = src, dir_name = file.path(x$package_dir, "ppt/media"))
  xml_elt <- fortify_pml_images(x, xml_elt)

  doc <- as_xml_document(xml_elt)

  node <- xml_find_first( doc, "p:spPr")
  off <- xml_child(node, "a:xfrm/a:off")
  xml_attr( off, "x") <- sprintf( "%.0f", left * 914400 )
  xml_attr( off, "y") <- sprintf( "%.0f", top * 914400 )
  if( rot != 0 ){
    xfrm_node <- xml_child(node, "a:xfrm")
    xml_attr( xfrm_node, "rot") <- sprintf( "%.0f", -rot * 60000 )
  }

  xmlslide <- slide$get()


  xml_add_child(xml_find_first(xmlslide, "p:cSld/p:spTree"), doc)
  slide$save()
  x$slide$update()
  x
}

#' @export
#' @rdname ph_with_table
#' @param left,top location of the new shape on the slide
#' @param width,height shape size in inches
#' @examples
#'
#' library(magrittr)
#'
#' doc <- read_pptx() %>%
#'   add_slide(layout = "Title and Content", master = "Office Theme") %>%
#'   ph_with_table_at(value = mtcars[1:6,],
#'     height = 4, width = 8, left = 4, top = 4,
#'     last_row = FALSE, last_column = FALSE, first_row = TRUE)
#'
#' print(doc, target = "ph_with_table2.pptx")
ph_with_table_at <- function( x, value, left, top, width, height,
                           first_row = TRUE, first_column = FALSE,
                           last_row = FALSE, last_column = FALSE ){
  stopifnot(is.data.frame(value))

  slide <- x$slide$get_slide(x$cursor)

  xml_elt <- table_shape(x = x, value = value, left = left, top = top, width = width, height = height,
                         first_row = first_row, first_column = first_column,
                         last_row = last_row, last_column = last_column )

  xml_add_child(xml_find_first(slide$get(), "//p:spTree"), as_xml_document(xml_elt))
  slide$save()
  x$slide$update()
  x
}

