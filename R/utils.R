XML_HEADER <- "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"

base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""

get_style_id <- function(x, style, type ){
  ref <- x$styles[x$styles$style_type==type, ]
  if(!style %in% ref$style_name){
    t_ <- shQuote(ref$style_name, type = "sh")
    t_ <- paste(t_, collapse = ", ")
    t_ <- paste0("c(", t_, ")")
    stop("could not match any style named ", shQuote(style, type = "sh"), " in ", t_, call. = FALSE)
  }
  ref$style_id[ref$style_name == style]
}

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
  xfrm <- map_df( nodeset, function(x) {
    ph <- xml_child(x, "p:nvSpPr/p:nvPr/p:ph")
    type <- xml_attr(ph, "type")
    if( is.na(type) )
      type <- "body"
    id <- xml_child(x, "p:nvSpPr/p:cNvPr")
    off <- xml_child(x, "p:spPr/a:xfrm/a:off")
    ext <- xml_child(x, "p:spPr/a:xfrm/a:ext")
    tibble( type = type,
            id = xml_attr(id, "id"),
            file = basename(file),
            offx = as.integer(xml_attr(off, "x")),
            offy = as.integer(xml_attr(off, "y")),
            cx = as.integer(xml_attr(ext, "cx")),
            cy = as.integer(xml_attr(ext, "cy")),
            name = name )
  })
}


#' @importFrom dplyr rename
#' @importFrom dplyr select
#' @importFrom dplyr left_join
#' @importFrom dplyr mutate
xfrmize <- function( x ){

  slide_xfrm <- x$slideLayouts$description()
  master_xfrm <- x$slideLayouts$get_master()$description() %>%
    rename(offx_ref = offx, offy_ref = offy, cx_ref = cx, cy_ref = cy) %>%
    select(-name, -id)

  slide_xfrm <- left_join(
    slide_xfrm,
    master_xfrm,
    by = c("master_file"="file", "type" = "type")
  )
  slide_xfrm <- slide_xfrm %>%
    mutate( offx = ifelse( !is.finite(offx), offx_ref, offx ),
            offy = ifelse( !is.finite(offy), offy_ref, offy ),
            cx = ifelse( !is.finite(cx), cx_ref, cx ),
            cy = ifelse( !is.finite(cy), cy_ref, cy ) ) %>%
    mutate( offx = offx / 914400,
            offy = offy / 914400,
            cx = cx / 914400,
            cy = cy / 914400 ) %>%
    select(-offx_ref, -offy_ref, -cx_ref, -cy_ref)

  slide_xfrm
}


