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
                   rotation = integer(0),
                   name = character(0),
                   fld_id = character(0),
                   fld_type = character(0)
                   ))
  }

  ph <- xml_child(nodeset, "p:nvSpPr/p:nvPr/p:ph")
  type <- xml_attr(ph, "type")
  type[is.na(type)] <- "body"
  id <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "id")
  label <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "name")

  off <- xml_child(nodeset, "p:spPr/a:xfrm/a:off")
  ext <- xml_child(nodeset, "p:spPr/a:xfrm/a:ext")
  rot <- xml_child(nodeset, "p:spPr/a:xfrm")

  fld_id <- xml_attr(xml_child(nodeset, "/p:txBody/a:p/a:fld"), "id")
  fld_type <- xml_attr(xml_child(nodeset, "/p:txBody/a:p/a:fld"), "type")

  data.frame(stringsAsFactors = FALSE, type = type, id = id,
          ph_label = label,
          ph = as.character(ph),
          file = basename(file),
          offx = as.integer(xml_attr(off, "x")),
          offy = as.integer(xml_attr(off, "y")),
          cx = as.integer(xml_attr(ext, "cx")),
          cy = as.integer(xml_attr(ext, "cy")),
          rotation = as.integer(xml_attr(rot, "rot")),
          fld_id,
          fld_type,
          name = name )
}

fortify_master_xfrm <- function(master_xfrm){

  master_xfrm <- as.data.frame(master_xfrm)
  has_type <- grepl("type=", master_xfrm$ph)
  master_xfrm <- master_xfrm[has_type, ]
  master_xfrm <- master_xfrm[!duplicated(master_xfrm$type),]

  tmp_names <- names(master_xfrm)

  old_ <- c("offx", "offy", "cx", "cy", "fld_id", "fld_type", "name")
  new_ <- c("offx_ref", "offy_ref", "cx_ref", "cy_ref", "fld_id_ref", "fld_type_ref", "master_name")
  tmp_names[match(old_, tmp_names)] <- new_
  names(master_xfrm) <- tmp_names
  master_xfrm$id <- NULL
  master_xfrm$ph <- NULL
  master_xfrm$ph_label <- NULL
  master_xfrm$rotation <- NULL

  master_xfrm
}

xfrmize <- function( slide_xfrm, master_xfrm ){
  x <- as.data.frame( slide_xfrm )

  master_ref <- unique( data.frame(file = master_xfrm$file,
                                     master_name = master_xfrm$name,
                                     stringsAsFactors = FALSE ) )
  master_xfrm <- fortify_master_xfrm(master_xfrm)

  slide_key_id <- paste0(x$master_file, x$type)
  master_key_id <- paste0(master_xfrm$file, master_xfrm$type)

  slide_xfrm_no_match <- x[!slide_key_id %in% master_key_id, ]
  slide_xfrm_no_match <- merge(slide_xfrm_no_match,
                               master_ref, by.x = "master_file", by.y = "file",
                               all.x = TRUE, all.y = FALSE)

  x <- merge(x, master_xfrm,
                      by.x = c("master_file", "type"),
                      by.y = c("file", "type"),
                      all = FALSE)
  x$offx <- ifelse( !is.finite(x$offx), x$offx_ref, x$offx )
  x$offy <- ifelse( !is.finite(x$offy), x$offy_ref, x$offy )
  x$cx <- ifelse( !is.finite(x$cx), x$cx_ref, x$cx )
  x$cy <- ifelse( !is.finite(x$cy), x$cy_ref, x$cy )
  x$offx_ref <- NULL
  x$offy_ref <- NULL
  x$cx_ref <- NULL
  x$cy_ref <- NULL
  x$fld_id_ref <- NULL
  x$fld_type_ref <- NULL

  x <- rbind(x, slide_xfrm_no_match, stringsAsFactors = FALSE)
  x[
    !is.na( x$offx ) &
      !is.na( x$offy ) &
      !is.na( x$cx ) &
      !is.na( x$cy ),]
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
    if( is.character(x) ) x
    else if( is.factor(x) ) as.character(x)
    else gsub("(^ | $)+", "", format(x))
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
  x <- Filter(function(x) nrow(x)>0, list_df)
  x <- lapply(x, function(x, col) {
    x[, setdiff(col, colnames(x))] <- NA
    x
  }, col = col)
  do.call(rbind, x)
}

set_row_span <- function( row_details ){
  row_details$first[!row_details$first & !row_details$row_merge] <- TRUE
  row_details$row_merge <- NULL

  row_details <- split(row_details, row_details$cell_id)

  row_details <- mapply(function(dat){
    rowspan_values_at_breaks <- rle(cumsum(dat$first))$lengths
    rowspan_pos_at_breaks <- which(dat$first)
    dat$row_span <- 0L
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

correct_id <- function(doc, int_id){
  all_uid <- xml_find_all(doc, "//*[@id]")
  for(z in seq_along(all_uid) ){
    if(!grepl("[^0-9]", xml_attr(all_uid[[z]], "id"))){
      xml_attr(all_uid[[z]], "id") <- int_id
      int_id <- int_id + 1
    }
  }
  int_id
}





check_bookmark_id <- function(bkm){
  if(!is.null(bkm)){
    invalid_bkm <- is.character(bkm) &&
      length(bkm) == 1 &&
      nchar(bkm) > 0 &&
      grepl("[^:[:alnum:]_-]+", bkm)
    if(invalid_bkm){
      stop("bkm [", bkm, "] should only contain alphanumeric characters, ':', '-' and '_'.", call. = FALSE)
    }
  }
  bkm
}

is_windows <- function() {
  "windows" %in% .Platform$OS.type
}

is_office_doc_edited <- function(file) {

  file_name <- sub("(.*)\\.(pptx|docx)$", "\\1", basename(file))

  if (nchar(file_name) < 7 || endsWith(file, ".pptx")) {
    edit_name <- paste0("~$", basename(file))
  } else if (nchar(file_name) < 8) {
    edit_name <- paste0("~$", substring(basename(file), 2))
  } else {
    edit_name <- paste0("~$", substring(basename(file), 3))
  }

  file.exists(file.path(dirname(file), edit_name))
}

.url_special_chars <- list(
  `&` = '&amp;',
  `<` = '&lt;',
  `>` = '&gt;',
  `'` = '&#39;',
  `"` = '&quot;',
  ` ` = "&nbsp;"
)
officer_url_encode <- function(x) {
  for (chr in names(.url_special_chars)) {
    x <- gsub(chr, .url_special_chars[[chr]], x, fixed = TRUE, useBytes = TRUE)
  }
  Encoding(x) <- "UTF-8"
  x
}
officer_url_decode <- function(x) {
  for (chr in rev(names(.url_special_chars))) {
    x <- gsub(.url_special_chars[[chr]], chr, x, fixed = TRUE, useBytes = TRUE)
  }
  Encoding(x) <- "UTF-8"
  x
}
# htmlEscapeCopy ----
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
