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

pml_shape_str <- function(str, ph, offx, offy, cx, cy, ...) {

  sp_pr <- sprintf("<p:spPr><a:xfrm><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm></p:spPr>", offx, offy, cx, cy)
  # sp_pr <- "<p:spPr/>"
  nv_sp_pr <- "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>"
  nv_sp_pr <- sprintf( nv_sp_pr, ifelse(!is.na(ph), ph, "") )
  paste0( pml_with_ns("p:sp"),
                  nv_sp_pr, sp_pr,
                  "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>",
                  htmlEscape(str),
                  "</a:t></a:r></a:p></p:txBody></p:sp>"
                  )
}


pml_shape_par <- function(str, ph, offx, offy, cx, cy, ...) {

  sp_pr <- sprintf("<p:spPr><a:xfrm><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm></p:spPr>", offx, offy, cx, cy)

  nv_sp_pr <- "<p:nvSpPr><p:cNvPr id=\"\" name=\"\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>"
  nv_sp_pr <- sprintf( nv_sp_pr, ifelse(!is.na(ph), ph, "") )
  paste0( pml_with_ns("p:sp"),
          nv_sp_pr, sp_pr,
          "<p:txBody><a:bodyPr/><a:lstStyle/>",
          str,
          "</p:txBody></p:sp>"
  )
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
    return(data.frame(stringsAsFactors = FALSE, type = character(0),
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

  data.frame(stringsAsFactors = FALSE, type = type, id = id,
          ph_label = label,
          ph = as.character(ph),
          file = basename(file),
          offx = as.integer(xml_attr(off, "x")),
          offy = as.integer(xml_attr(off, "y")),
          cx = as.integer(xml_attr(ext, "cx")),
          cy = as.integer(xml_attr(ext, "cy")),
          name = name )
}


fortify_master_xfrm <- function(master_xfrm){
  master_xfrm <- as.data.frame(master_xfrm)
  has_type <- grepl("type=", master_xfrm$ph)
  master_xfrm <- master_xfrm[has_type, ]
  master_xfrm <- master_xfrm[!duplicated(master_xfrm$type),]

  tmp_names <- names(master_xfrm)

  old_ <- c("offx", "offy", "cx", "cy", "name")
  new_ <- c("offx_ref", "offy_ref", "cx_ref", "cy_ref", "master_name")
  tmp_names[match(old_, tmp_names)] <- new_
  names(master_xfrm) <- tmp_names
  master_xfrm$id <- NULL
  master_xfrm$ph <- NULL
  master_xfrm$ph_label <- NULL

  master_xfrm
}

xfrmize <- function( slide_xfrm, master_xfrm ){

  slide_xfrm <- as.data.frame( slide_xfrm )

  master_ref <- unique( data.frame(file = master_xfrm$file,
                                     master_name = master_xfrm$name,
                                     stringsAsFactors = FALSE ) )
  master_xfrm <- fortify_master_xfrm(master_xfrm)

  slide_key_id <- paste0(slide_xfrm$master_file, slide_xfrm$type)
  master_key_id <- paste0(master_xfrm$file, master_xfrm$type)

  slide_xfrm_no_match <- slide_xfrm[!slide_key_id %in% master_key_id, ]
  slide_xfrm_no_match <- merge(slide_xfrm_no_match,
                               master_ref, by.x = "master_file", by.y = "file",
                               all.x = TRUE, all.y = FALSE)

  slide_xfrm <- merge(slide_xfrm, master_xfrm,
                      by.x = c("master_file", "type"),
                      by.y = c("file", "type"),
                      all = FALSE)
  slide_xfrm$offx <- ifelse( !is.finite(slide_xfrm$offx), slide_xfrm$offx_ref, slide_xfrm$offx )
  slide_xfrm$offy <- ifelse( !is.finite(slide_xfrm$offy), slide_xfrm$offy_ref, slide_xfrm$offy )
  slide_xfrm$cx <- ifelse( !is.finite(slide_xfrm$cx), slide_xfrm$cx_ref, slide_xfrm$cx )
  slide_xfrm$cy <- ifelse( !is.finite(slide_xfrm$cy), slide_xfrm$cy_ref, slide_xfrm$cy )
  slide_xfrm$offx_ref <- NULL
  slide_xfrm$offy_ref <- NULL
  slide_xfrm$cx_ref <- NULL
  slide_xfrm$cy_ref <- NULL

  slide_xfrm <- rbind(slide_xfrm, slide_xfrm_no_match, stringsAsFactors = FALSE)
  slide_xfrm[
    !is.na( slide_xfrm$offx ) &
      !is.na( slide_xfrm$offy ) &
      !is.na( slide_xfrm$cx ) &
      !is.na( slide_xfrm$cy ),]
}


set_xfrm_attr <- function( node, offx, offy, cx, cy ){
  off <- xml_child(node, "p:xfrm/a:off")
  ext <- xml_child(node, "p:xfrm/a:ext")

  xml_attr( off, "x") <- sprintf( "%.0f", offx )
  xml_attr( off, "y") <- sprintf( "%.0f", offy )
  xml_attr( ext, "cx") <- sprintf( "%.0f", cx )
  xml_attr( ext, "cy") <- sprintf( "%.0f", cy )

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
  data.frame(stringsAsFactors = FALSE, name = names_, type = col_types_, value = vals, theme = theme)
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
  names(x) <- htmlEscape(names(x))
  x <- lapply(x, function( x ) {
    if( is.character(x) ) htmlEscape(x)
    else if( is.factor(x) ) htmlEscape(as.character(x))
    else gsub("(^ | $)+", "", htmlEscape(format(x)))
  })
  data.frame(x, stringsAsFactors = FALSE, check.names = FALSE)
}


xpath_content_selector <- "*[self::p:cxnSp or self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"

as_xpath_content_sel <- function(prefix){
  paste0(prefix, xpath_content_selector)
}


between <- function(x, left, right ){
  x >= left & x <= right
}



simple_lag <- function( x, default=0 ){
  c(default, x[-length(x)])
}

rbind.match.columns <- function(list_df) {
  col <- unique(unlist(sapply(list_df, names)))

  list_df <- lapply(list_df, function(x, col) {
    x[, setdiff(col, names(x))] <- NA
    x
  }, col = col)
  do.call(rbind, list_df)
}

set_row_span <- function( row_details ){
  row_details$first[!row_details$first & !row_details$row_merge] <- TRUE
  row_details$row_merge <- NULL

  row_details <- split(row_details, row_details$cell_id)

  row_details <- mapply(function(dat){
    rowspan_values_at_breaks <- rle(cumsum(dat$first))$lengths
    rowspan_pos_at_breaks <- which(dat$first)
    dat$row_span <- 0
    dat$row_span[rowspan_pos_at_breaks] <- rowspan_values_at_breaks
    dat
  }, row_details, SIMPLIFY = FALSE)
  row_details <- rbind.match.columns(row_details)
  row_details$first <- NULL
  row_details
}


is_scalar_character <- function( x ) {
  is.character(x) && length(x) == 1
}
is_scalar_logical <- function( x ) {
  is.logical(x) && length(x) == 1
}
