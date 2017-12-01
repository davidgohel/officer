unfold_row_wml <- function(node, row_id){
  is_header_1 <- !inherits(xml_child(node, "w:trPr/w:tblHeader"), "xml_missing")
  is_header_2 <- !inherits(xml_child(node, "w:trPr/w:cnfStyle[@w:firstRow='1']"), "xml_missing")
  is_header <- is_header_1 | is_header_2
  children_ <- xml_children(node)
  cell_nodes <- children_[sapply(children_, function(x) xml_name(x)=="tc" )]

  txt <- sapply(cell_nodes, xml_text)
  col_span <- sapply(cell_nodes, function(x) {
    gs <- xml_child(x, "w:tcPr/w:gridSpan")
    as.integer(xml_attr(gs, "val"))
  })
  col_span[is.na(col_span)] <- 1

  row_span <- lapply(cell_nodes, function(x) {
    node_ <- xml_child(x, "w:tcPr/w:vMerge")
    data.frame(row_merge = !inherits(node_, "xml_missing"),
           first = xml_attr(node_, "val") %in% "restart",
           row_span = 1,
           stringsAsFactors = FALSE
    )
  })
  row_span <- rbind.match.columns(row_span)
  txt[row_span$row_merge & !row_span$first] <- NA

  out <- data.frame(row_id = row_id, is_header = is_header,
                cell_id = 1 + simple_lag( cumsum(col_span), default=0 ),
                text = txt, col_span = col_span, stringsAsFactors = FALSE)
  out <- cbind(out, row_span)

  out_add_ <- out[out$col_span > 1, ]

  out_add_ <- mapply(function(row_id, is_header, cell_id, text, col_span, row_merge, first, row_span){
    reps_ <- col_span - 1
    row_id_ <- rep( row_id, reps_)
    is_header_ <- rep( is_header, reps_)
    cell_id_ <- rep( cell_id, reps_)
    text_ <- rep( NA, reps_)
    col_span_ <- rep( 0, reps_)
    row_merge_ <- rep( row_merge, reps_)
    first_ <- rep( first, reps_)
    row_span_ <- rep( row_span, reps_)
    out <- data.frame(row_id = row_id_, is_header = is_header_, cell_id = cell_id_,
                      text = text_, col_span = col_span_, row_merge = row_merge_,
                      first = first_, row_span = row_span_,
                      stringsAsFactors = FALSE)
    out$cell_id <- seq_len(reps_) + cell_id
    out
  }, out_add_$row_id, out_add_$is_header, out_add_$cell_id, out_add_$text,
  out_add_$col_span, out_add_$row_merge, out_add_$first, out_add_$row_span,
  SIMPLIFY = FALSE)
  if( length(out_add_) > 0 ){
    out_add_ <- rbind.match.columns(out_add_)
    out <- rbind(out, out_add_)
  }
  out[order(out$cell_id),]
}


docxtable_as_tibble <- function( node, styles ){
  xpath_ <- paste0( xml_path(node), "/w:tr")
  rows <- xml_find_all(node, xpath_)
  if( length(rows) < 1 ) return(NULL)

  row_details <- mapply(unfold_row_wml, rows, seq_along(rows), SIMPLIFY = FALSE)
  row_details <- rbind.match.columns(row_details)
  row_details <- set_row_span(row_details)

  style_node <- xml_child(node, "w:tblPr/w:tblStyle")
  if( inherits(style_node, "xml_missing") ){
    style_name <- NA
  } else {
    style_id <- xml_attr( style_node, "val")
    row_details$style_name <- rep( styles$style_name[styles$style_id %in% style_id], nrow(row_details))
  }
  row_details$content_type <- rep("table cell", nrow(row_details) )
  row_details$text[row_details$col_span < 1 | row_details$row_span < 1] <- NA_character_
  row_details
}

par_as_tibble <- function(node, styles){

  style_node <- xml_child(node, "w:pPr/w:pStyle")
  if( inherits(style_node, "xml_missing") ){
    style_name <- NA
  } else {
    style_id <- xml_attr( style_node, "val")
    style_name <- styles$style_name[styles$style_id %in% style_id]
  }

  par_data <- data.frame(
    level = as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:ilvl"), "val")) + 1,
    num_id = as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:numId"), "val")),
    text = xml_text(node), style_name = style_name,
    stringsAsFactors = FALSE )

  par_data$content_type <- rep("paragraph", nrow(par_data) )
  par_data
}

node_content <- function(node, x){
  node_name <- xml_name(node)
  switch(node_name,
         p = par_as_tibble(node, styles_info(x)),
         tbl = docxtable_as_tibble(node, styles_info(x)),
         NULL)
}


#' @title get Word content in a tidy format
#' @description read content of a Word document and
#' return a tidy dataset representing the document.
#' @param x an rdocx object
#' @examples
#' example_pptx <- system.file(package = "officer",
#'   "doc_examples/example.docx")
#' doc <- read_docx(example_pptx)
#' docx_summary(doc)
#' @export
docx_summary <- function( x ){

  all_nodes <- xml_find_all(x$doc_obj$get(),"/w:document/w:body/*[self::w:p or self::w:tbl]")
  data <- lapply( all_nodes, node_content, x = x )

  data <- mapply(function(x, id) {
        x$doc_index <- id
        x
      }, data, seq_along(data), SIMPLIFY = FALSE)
  data <- rbind.match.columns(data)

  colnames <- c("doc_index", "content_type", "style_name", "text",
    "level", "num_id", "row_id", "is_header", "cell_id",
    "col_span", "row_span")
  colnames <- intersect(colnames, names(data))
  data[, colnames]
}

