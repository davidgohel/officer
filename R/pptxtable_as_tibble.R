#' @importFrom dplyr bind_cols lag rowwise do
unfold_row_pml <- function(node, row_id){

  children_ <- xml_children(node)
  cell_nodes <- children_[map_lgl(children_, function(x) xml_name(x)=="tc" )]

  txt <- map_chr(cell_nodes, xml_text)

  col_span <- map_int(cell_nodes, function(x) {
    as.integer(xml_attr(x, "gridSpan"))
  })
  h_merge <- map_int(cell_nodes, function(x) {
    as.integer(xml_attr(x, "hMerge"))
  }) %in% c(1)
  col_span[is.na(col_span)] <- 1
  col_span[h_merge] <- 0

  row_span <- map_int(cell_nodes, function(x) {
    as.integer(xml_attr(x, "rowSpan"))
  })
  row_span[is.na(row_span)] <- 1

  out <- tibble(row_id = row_id, cell_id = seq_along(cell_nodes),
                txt = txt, col_span = col_span, row_span = row_span)

  out
}

globalVariables(c("."))

#' @importFrom purrr map2_df
pptxtable_as_tibble <- function( node ){
  xpath_ <- paste0( xml_path(node), "/a:graphic/a:graphicData/a:tbl/a:tr")
  rows <- xml_find_all(node, xpath_)
  if( length(rows) < 1 ) return(NULL)

  row_details <- map2_df( rows, seq_along(rows), unfold_row_pml )
  rep_rows <- rowwise(row_details[row_details$row_span> 1,])
  rep_rows <- do(rep_rows, {
    data <- .
    row_id <- data$row_id + seq_len(data$row_span)
    row_span <- rep(0, data$row_span)
    cell_id <- rep(data$cell_id, data$row_span)
    col_span <- rep(data$col_span, data$row_span)
    data.frame(row_id = row_id, row_span = row_span, col_span = col_span, cell_id = cell_id)
  }) %>% ungroup()

  if( nrow(rep_rows) > 0 ){
    row_details <- anti_join(row_details, rep_rows, by = c("row_id", "cell_id") )
    row_details <- bind_rows(row_details, rep_rows)
  }
  arrange_(row_details, .dots = c("row_id", "cell_id") )

}


pptx_par_as_tibble <- function(node){
  xpath_ <- paste0( xml_path( node ), "/p:txBody/a:p")
  p_nodes <- xml_find_all(node, xpath_ )
  tibble( paragraph = xml_text(p_nodes) )
}

#' @importFrom tools file_ext
#' @importFrom png readPNG
#' @importFrom jpeg readJPEG
#' @importFrom bmp read.bmp
rasterize_img <- function( path ){
  ext <- tools::file_ext(path)
  switch(ext,
         png = readPNG(path),
         jpeg = readJPEG(path),
         jpg = readJPEG(path),
         bmp = read.bmp(path),
         NULL
         )

}

pptx_img_as_tibble  <- function(node, img_src ){
  blip <- xml_child(node, "/p:blipFill/a:blip" )
  img_id <- xml_attr(blip, "embed")
  rasters <- rasterize_img( img_src[img_id] )
  rasters
}

