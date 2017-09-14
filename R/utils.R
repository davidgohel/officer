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
  sprintf(str_, format(style, type = "pml"), htmlEscape(str))
}

pml_shape_str <- function(str, ph) {
  str_ <- paste0( pml_with_ns("p:sp"),
                  "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>",
                  "<p:spPr/>",
                  "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>"
                  )
  sprintf( str_, ph, htmlEscape(str) )
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
                   ph_label = character(0),
                   ph = character(0),
                   file = character(0),
                   offx = integer(0),
                   offy = integer(0),
                   cx = integer(0),
                   cy = integer(0),
                   name = character(0) ))
  }

  ph <- xml_child(nodeset, "p:nvSpPr/p:nvPr/p:ph")
  type <- xml_attr(ph, "type")
  type[is.na(type)] <- "body"
  id <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "id")
  label <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "name")
  off <- xml_child(nodeset, "p:spPr/a:xfrm/a:off")
  ext <- xml_child(nodeset, "p:spPr/a:xfrm/a:ext")
  tibble( type = type, id = id,
          ph_label = label,
          ph = as.character(ph),
          file = basename(file),
          offx = as.integer(xml_attr(off, "x")),
          offy = as.integer(xml_attr(off, "y")),
          cx = as.integer(xml_attr(ext, "cx")),
          cy = as.integer(xml_attr(ext, "cy")),
          name = name )
}


#' @importFrom dplyr left_join anti_join bind_rows distinct
xfrmize <- function( slide_xfrm, master_xfrm ){
  slide_xfrm <- as.data.frame( slide_xfrm )
  master_xfrm <- as.data.frame(master_xfrm)

  master_ref <- distinct( data.frame(file = master_xfrm$file,
                                     master_name = master_xfrm$name,
                                     stringsAsFactors = FALSE ) )
  tmp_names <- names(master_xfrm)
  old_ <- c("offx", "offy", "cx", "cy", "name")
  new_ <- c("offx_ref", "offy_ref", "cx_ref", "cy_ref", "master_name")
  tmp_names[match(old_, tmp_names)] <- new_
  names(master_xfrm) <- tmp_names
  master_xfrm$id <- NULL
  master_xfrm$ph <- NULL

  slide_key_id <- paste0(slide_xfrm$master_file, slide_xfrm$type)
  master_key_id <- paste0(master_xfrm$file, master_xfrm$type)

  slide_xfrm_no_match <- slide_xfrm[!slide_key_id %in% master_key_id, ] %>%
    inner_join( master_ref, by = c("master_file"="file") )

  slide_xfrm <- inner_join(
    slide_xfrm,
    master_xfrm,
    by = c("master_file"="file", "type" = "type")
  )

  slide_xfrm$offx <- ifelse( !is.finite(slide_xfrm$offx), slide_xfrm$offx_ref, slide_xfrm$offx ) / 914400
  slide_xfrm$offy <- ifelse( !is.finite(slide_xfrm$offy), slide_xfrm$offy_ref, slide_xfrm$offy ) / 914400
  slide_xfrm$cx <- ifelse( !is.finite(slide_xfrm$cx), slide_xfrm$cx_ref, slide_xfrm$cx ) / 914400
  slide_xfrm$cy <- ifelse( !is.finite(slide_xfrm$cy), slide_xfrm$cy_ref, slide_xfrm$cy ) / 914400
  slide_xfrm$offx_ref <- NULL
  slide_xfrm$offy_ref <- NULL
  slide_xfrm$cx_ref <- NULL
  slide_xfrm$cy_ref <- NULL
  bind_rows(slide_xfrm, slide_xfrm_no_match)
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
    filter_v <- shape_index_data$type == type & shape_index_data$id == id_chr
    shape_index_data <- shape_index_data[filter_v,]
  } else if( is.null(type) && !is.null(id_chr) ){
    filter_v <- shape_index_data$id == id_chr
    shape_index_data <- shape_index_data[filter_v,]
  } else if( !is.null(type) ){
    filter_v <- shape_index_data$type == type
    shape_index_data <- shape_index_data[filter_v,]
  }

  if( nrow(shape_index_data) < 1 )
    stop("selection does not match any row in slide_summary. Use function slide_summary.", call. = FALSE)
  else if( nrow(shape_index_data) > 1 )
    stop("selection does match more than a single row in slide_summary. Use function slide_summary.", call. = FALSE)

  shape_index_data$shape_id
}


characterise_df <- function(x){
  x <- lapply(x, function( x ) {
    if( is.character(x) ) htmlEscape(x)
    else if( is.factor(x) ) htmlEscape(as.character(x))
    else gsub("(^ | $)+", "", htmlEscape(format(x)))
  })
  data.frame(x, stringsAsFactors = FALSE)
}

section_dimensions <- function(node){
  section_obj <- as_list(node)

  landscape <- FALSE
  if( !is.null(attr(section_obj$pgSz, "orient")) && attr(section_obj$pgSz, "orient") == "landscape" ){
    landscape <- TRUE
  }

  h_ref <- as.integer(attr(section_obj$pgSz, "h"))
  w_ref <- as.integer(attr(section_obj$pgSz, "w"))

  mar_t <- as.integer(attr(section_obj$pgMar, "top"))
  mar_b <- as.integer(attr(section_obj$pgMar, "bottom"))
  mar_r <- as.integer(attr(section_obj$pgMar, "right"))
  mar_l <- as.integer(attr(section_obj$pgMar, "left"))
  mar_h <- as.integer(attr(section_obj$pgMar, "header"))
  mar_f <- as.integer(attr(section_obj$pgMar, "footer"))

  list( page = c("width" = w_ref, "height" = h_ref),
        landscape = landscape,
        margins = c(top = mar_t, bottom = mar_b,
                    left = mar_l, right = mar_r,
                    header = mar_h, footer = mar_f) )

}
