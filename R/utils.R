wml_with_ns <- function(x){
  base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
  sprintf("<%s %s>", x, base_ns)
}

pml_with_ns <- function(x){
  base_ns <- "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\""
  sprintf("<%s %s>", x, base_ns)
}

pml_run_str <- function(str, style) {
  str_ <- paste0( pml_with_ns("a:r"), "%s<a:t>%s</a:t></a:r>" )
  sprintf(str_, format(style, type = "pml"), str)
}

pml_shape_str <- function(str, ph) {
  str_ <- paste0( pml_with_ns("p:sp"),
                  "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>",
                  "<p:spPr/>",
                  "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>"
                  )
  sprintf( str_, ph, str )
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



read_theme_colors <- function(doc, theme){

  nodes <- xml_find_all(doc, "//a:clrScheme/*")

  names_ <- xml_name(nodes)
  col_types_ <- xml_name(xml_children(nodes) )
  vals <- xml_attr(xml_children(nodes), "val")
  last_colors_ <- xml_attr(xml_children(nodes), "lastClr")
  vals <- ifelse(col_types_ == "srgbClr", paste0("#", vals), paste0("#", last_colors_) )
  tibble(name = names_, type = col_types_, value = vals, theme = theme)
}


get_shape_id <- function(x, type = NULL, id_chr = NULL ){
  shape_index_data <- slide_summary(x)
  shape_index_data$shape_id <- seq_len(nrow(shape_index_data))

  if( !is.null(type) && !is.null(id_chr) ){
    filter_criteria <- interp(~ type == tp & id == index, tp = type, index = id_chr)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else if( is.null(type) && !is.null(id_chr) ){
    filter_criteria <- interp(~ id == index, index = id_chr)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else if( !is.null(type) && is.null(id_chr) ){
    filter_criteria <- interp(~ type == tp, tp = type)
    shape_index_data <- filter_(shape_index_data, filter_criteria)
  } else {
    filter_criteria <- interp(~ type == tp, tp = type)
    shape_index_data <- shape_index_data[nrow(shape_index_data), ]
  }

  if( nrow(shape_index_data) < 1 )
    stop("selection does not match any row in slide_summary. Use function slide_summary.", call. = FALSE)
  else if( nrow(shape_index_data) > 1 )
    stop("selection does match more than a single row in slide_summary. Use function slide_summary.", call. = FALSE)

  shape_index_data$shape_id
}


