#' @importFrom dplyr bind_cols lag rowwise do
unfold_row_xml <- function(node, row_id){
  is_header_1 <- !inherits(xml_child(node, "w:trPr/w:tblHeader"), "xml_missing")
  is_header_2 <- !inherits(xml_child(node, "w:trPr/w:cnfStyle[@w:firstRow='1']"), "xml_missing")
  is_header <- is_header_1 | is_header_2
  children_ <- xml_children(node)
  cell_nodes <- children_[map_lgl(children_, function(x) xml_name(x)=="tc" )]

  txt <- map_chr(cell_nodes, xml_text)
  col_span <- map_int(cell_nodes, function(x) {
    gs <- xml_child(x, "w:tcPr/w:gridSpan")
    as.integer(xml_attr(gs, "val"))
  })
  col_span[is.na(col_span)] <- 1

  row_span <- map_df(cell_nodes, function(x) {
    node_ <- xml_child(x, "w:tcPr/w:vMerge")
    tibble(row_merge = !inherits(node_, "xml_missing"),
           first = xml_attr(node_, "val") %in% "restart",
           row_span = 1
    )
  })

  txt[row_span$row_merge & !row_span$first] <- NA

  out <- tibble(row_id = row_id, is_header = is_header,
                cell_id = 1 + dplyr::lag( cumsum(col_span), default=0 ),
                text = txt, col_span = col_span) %>%
    bind_cols(row_span)


  out_add <- rowwise(out[out$col_span > 1, ])
  out_add <- do(out_add, {
      row_data <- .
      reps_ <- row_data$col_span - 1
      out <- map_df( seq_len(reps_), function(x, df) {
        out <- df
        out$col_span <- 0
        out$text <- NA
        out
      }, row_data)
      out$cell_id <- seq_len(reps_) + row_data$cell_id
      out
    })
  out <- bind_rows(out, out_add)
  out[order(out$cell_id),]
}


#' @importFrom purrr map2_df
docxtable_as_tibble <- function( node, styles ){
  xpath_ <- paste0( xml_path(node), "/w:tr")
  rows <- xml_find_all(node, xpath_)
  if( length(rows) < 1 ) return(NULL)

  row_details <- map2_df( rows, seq_along(rows), unfold_row_xml )
  out <- split(row_details, row_details$cell_id)
  out <- map_df(out, function(dat){
    rle_ <- rle(dat$row_merge)
    new_vals <- rle_$lengths[rle_$values]
    dat[dat$row_merge, "row_span"] <- 0
    dat[dat$row_merge & dat$first, "row_span"] <- new_vals
    dat$row_merge <- NULL
    dat$first <- NULL
    dat
  })

  style_node <- xml_child(node, "w:tblPr/w:tblStyle")
  if( inherits(style_node, "xml_missing") ){
    style_name <- NA
  } else {
    style_id <- xml_attr( style_node, "val")
    out$style_name <- rep( styles$style_name[styles$style_id %in% style_id], nrow(out))
  }

  add_column(out, content_type = "table cell")
}

par_as_tibble <- function(node, styles){

  style_node <- xml_child(node, "w:pPr/w:pStyle")
  if( inherits(style_node, "xml_missing") ){
    style_name <- NA
  } else {
    style_id <- xml_attr( style_node, "val")
    style_name <- styles$style_name[styles$style_id %in% style_id]
  }

  par_data <- tibble(
    level = as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:ilvl"), "val")) + 1,
    num_id = as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:numId"), "val")),
    text = xml_text(node), style_name = style_name )

  add_column(par_data, content_type = "paragraph")
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
#' @importFrom purrr map map2_df
#' @importFrom tibble add_column
docx_summary <- function( x ){

  all_nodes <- xml_find_all(x$doc_obj$get(),"/w:document/w:body/*[self::w:p or self::w:tbl]")
  data <- map(all_nodes, node_content, x)
  data <- map2_df( data, seq_along(data), function(x, id) {
    add_column(x, doc_index = id)
  })
  colnames <- c("doc_index", "content_type", "style_name", "text",
  "level", "num_id", "row_id", "is_header", "cell_id",
  "col_span", "row_span")
  colnames <- intersect(colnames, names(data))
  data[, colnames]
}

