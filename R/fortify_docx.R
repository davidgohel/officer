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
                txt = txt, col_span = col_span) %>%
    bind_cols(row_span)


  out_add <- rowwise(out[out$col_span > 1, ])
  out_add <- do(out_add, {
      row_data <- .
      reps_ <- row_data$col_span - 1
      out <- map_df( seq_len(reps_), function(x, df) {
        out <- df
        out$col_span <- 0
        out$txt <- NA
        out
      }, row_data)
      out$cell_id <- seq_len(reps_) + row_data$cell_id
      out
    })
  out <- bind_rows(out, out_add)
  arrange_(out, .dots = "cell_id")
}

globalVariables(c("."))

#' @importFrom purrr map2_df
docxtable_as_tibble <- function(node, styles ){
  xpath_ <- paste0( xml_path(node), "/w:tr")
  rows <- xml_find_all(node, xpath_)
  if( length(rows) < 1 ) return(NULL)

  row_details <- map2_df( rows, seq_along(rows), unfold_row_xml )
  out <- split(row_details, row_details$cell_id)
  map_df(out, function(dat){
    rle_ <- rle(dat$row_merge)
    new_vals <- rle_$lengths[rle_$values]
    dat[dat$row_merge, "row_span"] <- 0
    dat[dat$row_merge & dat$first, "row_span"] <- new_vals
    select_( dat, "-row_merge", "-first" )
  })
}

#' @importFrom purrr map2_df
docxtable_desc <- function(node, styles ){
  style_id <- xml_attr( xml_child(node, "w:tblPr/w:tblStyle"), "val")
  out <- tibble( content = "table", table_data = list(docxtable_as_tibble(node, styles )), style_id = style_id )
  styles <- select_( styles, "style_name", "style_id" )
  out <- left_join(out, styles, by = "style_id")
  select_( out, "-style_id")
}

par_as_tibble <- function(node, styles){
  txt <- xml_text(node)

  style_id <- xml_attr( xml_child(node, "w:pPr/w:pStyle"), "val")
  level <- as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:ilvl"), "val")) + 1
  num_id <- as.integer(xml_attr( xml_child(node, "w:pPr/w:numPr/w:numId"), "val"))

  out <- tibble( content = "paragraph", txt = txt, style_id = style_id, item_level = level, list_id = num_id )
  styles <- select_( styles, "style_name", "style_id" )
  out <- left_join(out, styles, by = "style_id")
  out <- select_( out, "-style_id")
  sect_data <- sect_as_tibble(xml_child(node, "w:pPr/w:sectPr"))
  bind_cols( out, sect_data )
}


sect_as_tibble <- function(node){

  if( inherits(node, "xml_missing") )
    return(NULL )

  sect_desc <- tibble(
    type = xml_attr( xml_child(node, "w:type"), "val"),
    width = as.integer( xml_attr( xml_child(node, "w:pgSz"), "w") ) / 12700,
    height = as.integer( xml_attr( xml_child(node, "w:pgSz"), "h") ) / 12700,
    top = as.integer( xml_attr( xml_child(node, "w:pgMar"), "top") ) / 12700,
    right = as.integer( xml_attr( xml_child(node, "w:pgMar"), "right") ) / 12700,
    bottom = as.integer( xml_attr( xml_child(node, "w:pgMar"), "bottom") ) / 12700,
    left = as.integer( xml_attr( xml_child(node, "w:pgMar"), "left") ) / 12700,
    header = as.integer( xml_attr( xml_child(node, "w:pgMar"), "header") ) / 12700,
    footer = as.integer( xml_attr( xml_child(node, "w:pgMar"), "footer") ) / 12700,
    gutter = as.integer( xml_attr( xml_child(node, "w:pgMar"), "gutter") ) / 12700,
    cols = as.integer( xml_attr( xml_child(node, "w:cols"), "num") ) )

  tibble(content = "section_break", section_data = list(sect_desc) )
}


#' @title tidy content for a docx object
#' @description read content of a Word document and
#' return a tidy dataset representing the document.
#' @param x an rdocx object
#' @examples
#' library(dplyr)
#' # dummy document ---
#' doc <- read_docx() %>%
#'   body_add_par("A title", style = "heading 1") %>%
#'   body_add_par("Hello world!", style = "Normal") %>%
#'   body_add_par("iris table", style = "centered") %>%
#'   body_add_table(iris, style = "table_template") %>%
#'   body_add_par("mtcars table", style = "centered") %>%
#'   body_add_table(mtcars, style = "table_template") %>%
#'   cursor_begin() %>% body_remove()
#'
#' # read the document in a tibble ---
#' my_doc_desc <- fortify_docx(doc)
#'
#' # access first table ---
#' my_doc_desc
#' all_tables <- my_doc_desc %>%
#'   filter(content=="table")
#' all_tables[[1, "table_data"]]
#' all_tables[[1, "table_data"]] %>%
#'   filter(is_header)
#' @export
fortify_docx <- function( x ){

  all_nodes <- xml_find_all(x$doc_obj$get(),"/w:document/w:body/*")
  data <- map2_df(all_nodes, seq_along(all_nodes), function(node, id){
    node_name <- xml_name(node)
    switch(node_name,
           p = par_as_tibble(node, styles_info(x)),
           tbl = docxtable_desc(node, styles_info(x)),
           sectPr = sect_as_tibble(node),
           NULL)
  })
  data <- select_(data, "content", "txt", "style_name", "table_data", "section_data", "item_level", "list_id")
  data$doc_index = seq_len(nrow(data))
  data
}
