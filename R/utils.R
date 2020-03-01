#' @importFrom xml2 xml_attr<- xml_name<- xml_text<- as_list as_xml_document read_xml
#' write_xml xml_add_child xml_add_parent xml_add_sibling xml_attr xml_attrs xml_child
#' xml_children xml_find_all xml_find_first xml_length xml_missing xml_name xml_ns
#' xml_path xml_remove xml_replace xml_set_attr xml_set_attrs xml_text
#' @importFrom stats setNames


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

fortify_pml_images <- function(x, str){

  slide <- x$slide$get_slide(x$cursor)
  ref <- slide$rel_df()

  ref <- ref[ref$ext_src != "",]
  doc <- as_xml_document(str)
  for(id in seq_along(ref$ext_src) ){
    xpth <- paste0("//p:pic/p:blipFill/a:blip",
                   sprintf( "[contains(@r:embed,'%s')]", ref$ext_src[id]),
                   "")

    src_nodes <- xml_find_all(doc, xpth)
    xml_attr(src_nodes, "r:embed") <- ref$id[id]
  }
  as.character(doc)
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


read_theme_colors <- function(doc, theme){

  nodes <- xml_find_all(doc, "//a:clrScheme/*")

  names_ <- xml_name(nodes)
  col_types_ <- xml_name(xml_children(nodes) )
  vals <- xml_attr(xml_children(nodes), "val")
  last_colors_ <- xml_attr(xml_children(nodes), "lastClr")
  vals <- ifelse(col_types_ == "srgbClr", paste0("#", vals), paste0("#", last_colors_) )
  data.frame(stringsAsFactors = FALSE, name = names_, type = col_types_, value = vals, theme = theme)
}



characterise_df <- function(x){
  names(x) <- htmlEscapeCopy(names(x))
  x <- lapply(x, function( x ) {
    if( is.character(x) ) htmlEscapeCopy(x)
    else if( is.factor(x) ) htmlEscapeCopy(as.character(x))
    else gsub("(^ | $)+", "", htmlEscapeCopy(format(x)))
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

  col <- unique(unlist(lapply(list_df, colnames)))
  list_df <- Filter(function(x) nrow(x)>0, list_df)
  list_df <- lapply(list_df, function(x, col) {
    x[, setdiff(col, colnames(x))] <- NA
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


#' @importFrom grDevices col2rgb rgb
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


htmlEscapeCopy <- local({

  .htmlSpecials <- list(
    `&` = '&amp;',
    `<` = '&lt;',
    `>` = '&gt;'
  )
  .htmlSpecialsPattern <- paste(names(.htmlSpecials), collapse='|')
  .htmlSpecialsAttrib <- c(
    .htmlSpecials,
    `'` = '&#39;',
    `"` = '&quot;',
    `\r` = '&#13;',
    `\n` = '&#10;'
  )
  .htmlSpecialsPatternAttrib <- paste(names(.htmlSpecialsAttrib), collapse='|')
  function(text, attribute=FALSE) {
    pattern <- if(attribute)
      .htmlSpecialsPatternAttrib
    else
      .htmlSpecialsPattern
    text <- enc2utf8(as.character(text))
    # Short circuit in the common case that there's nothing to escape
    if (!any(grepl(pattern, text, useBytes = TRUE)))
      return(text)
    specials <- if(attribute)
      .htmlSpecialsAttrib
    else
      .htmlSpecials
    for (chr in names(specials)) {
      text <- gsub(chr, specials[[chr]], text, fixed = TRUE, useBytes = TRUE)
    }
    Encoding(text) <- "UTF-8"
    return(text)
  }
})





