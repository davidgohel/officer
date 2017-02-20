XML_HEADER <- "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"

base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""

is.color = function(x) {
  # http://stackoverflow.com/a/13290832/3315962
  out = sapply(x, function( x ) {
    tryCatch( is.matrix( col2rgb( x ) ), error = function( e ) F )
  })

  nout <- names(out)
  if( !is.null(nout) && any( is.na( nout ) ) )
    out[is.na( nout )] = FALSE

  out
}


attr_chunk <- function( x ){
  if( !is.null(x) && length( x ) > 0){
    attribs <- paste0(names(x), "=", shQuote(x, type = "cmd"), collapse = " " )
    attribs <- paste0(" ", attribs)
  } else attribs <- ""
  attribs
}

read_xfrm <- function(nodeset, file, name){
  if( length(nodeset) < 1 ){
    return(tibble( type = character(0),
                   id = character(0),
                   ph = character(0),
                   file = character(0),
                   offx = integer(0),
                   offy = integer(0),
                   cx = integer(0),
                   cy = integer(0),
                   name = character(0) ))
  }

  xfrm <- map_df( nodeset, function(x) {
    ph <- xml_child(x, "p:nvSpPr/p:nvPr/p:ph")
    type <- xml_attr(ph, "type")
    if( is.na(type) )
      type <- "body"
    id <- xml_child(x, "/p:cNvPr")
    off <- xml_child(x, "p:spPr/a:xfrm/a:off")
    ext <- xml_child(x, "p:spPr/a:xfrm/a:ext")
    tibble( type = type,
            id = xml_attr(id, "id"),
            ph = as.character(ph),
            file = basename(file),
            offx = as.integer(xml_attr(off, "x")),
            offy = as.integer(xml_attr(off, "y")),
            cx = as.integer(xml_attr(ext, "cx")),
            cy = as.integer(xml_attr(ext, "cy")),
            name = name )
  })
}


#' @importFrom dplyr left_join anti_join bind_rows distinct
#' @importFrom dplyr rename_ select_ mutate_
xfrmize <- function( slide_xfrm, master_xfrm ){

  master_ref <- master_xfrm %>%
    rename_( .dots = setNames( "name", "master_name")) %>%
    select_("file", "master_name") %>%
    distinct()

  master_xfrm <- master_xfrm %>%
    rename_( .dots = setNames( c("offx", "offy", "cx", "cy", "name"),
                               c("offx_ref", "offy_ref", "cx_ref", "cy_ref", "master_name"))) %>%
    select_("-id", "-ph")

  slide_xfrm_no_match <- anti_join(
    slide_xfrm,
    master_xfrm,
    by = c("master_file"="file", "type" = "type") ) %>% inner_join(
    master_ref,
    by = c("master_file"="file")
  )
  slide_xfrm <- inner_join(
    slide_xfrm,
    master_xfrm,
    by = c("master_file"="file", "type" = "type")
  )
  offx <- interp("ifelse( !is.finite(offx), offx_ref, offx )")
  offy <- interp("ifelse( !is.finite(offy), offy_ref, offy )")
  cx <- interp("ifelse( !is.finite(cx), cx_ref, cx )")
  cy <- interp("ifelse( !is.finite(cy), cy_ref, cy )")

  slide_xfrm <- slide_xfrm %>%
    mutate_( .dots = list(offx = offx, offy = offy, cx = cx, cy = cy ) ) %>%
    mutate_( .dots = list(offx = interp("offx / 914400"),
            offy = interp("offy / 914400"),
            cx = interp("cx / 914400"),
            cy = interp("cy / 914400") ) ) %>%
    select_("-offx_ref", "-offy_ref", "-cx_ref", "-cy_ref") %>%
    bind_rows(slide_xfrm_no_match)

  slide_xfrm
}


set_xfrm_attr <- function( node, offx, offy, cx, cy ){
  off <- xml_child(node, "p:xfrm/a:off")
  ext <- xml_child(node, "p:xfrm/a:ext")

  xml_attr( off, "x") <- sprintf( "%.0f", offx * 914400 )
  xml_attr( off, "y") <- sprintf( "%.0f", offy * 914400 )
  xml_attr( ext, "cx") <- sprintf( "%.0f", cx * 914400 )
  xml_attr( ext, "cy") <- sprintf( "%.0f", cy * 914400 )

  cnvpr <- xml_child(node, "*/p:cNvPr")
  xml_attr( cnvpr, "id") <- ""
  node
}

