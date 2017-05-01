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
                text = txt, col_span = col_span, row_span = row_span)

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
  tibble( text = xml_text(p_nodes) )
}



embed_img_raster  <- function(node, img_src ){
  blip <- xml_child(node, "/p:blipFill/a:blip" )
  img_id <- xml_attr(blip, "embed")

  file_ <- img_src[img_id]
  stopifnot(is.character(file_), length(file_) == 1)
  tibble(media_file = file.path( "ppt/media/", basename(file_) ) )
}

#' @export
#' @title Extract media from a document object
#' @description Extract files from an \code{rdocx} or \code{rpptx} object.
#' @param x an rpptx object or an rdocx object
#' @param path media path, should be a relative path
#' @param target target file
#' @examples
#' example_pptx <- system.file(package = "officer",
#'   "doc_examples/example.pptx")
#' doc <- read_pptx(example_pptx)
#' content <- pptx_summary(doc)
#' image_row <- content[content$content_type %in% "image", ]
#' media_file <- image_row$media_file
#' media_extract(doc, path = media_file, target = "extract.png")
media_extract <- function( x, path, target ){
  media <- file.path(x$package_dir, path )
  stopifnot(file.exists(media))
  file.copy(from = media, to = target)
}

#' @title get PowerPoint content in a tidy format
#' @description read content of a PowerPoint document and
#' return a tidy dataset representing the document.
#' @param x an rpptx object
#' @examples
#' example_pptx <- system.file(package = "officer",
#'   "doc_examples/example.pptx")
#' doc <- read_pptx(example_pptx)
#' pptx_summary(doc)
#' @export
pptx_summary <- function( x ){

  slide_index <- seq_len(length(x))

  map_df(slide_index, function(i, x){
    slide <- x$slide$get_slide(i)
    str = "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"
    nodes <- xml_find_all(slide$get(), str)
    data <- read_xfrm(nodes, file = "slide", name = "" )
    content <- map2_df(nodes, data$id, function(node, id, slide_id){
      is_table <- !inherits( xml_child(node, "/a:graphic/a:graphicData/a:tbl"), "xml_missing")
      is_par <- !inherits( xml_child(node, "/p:txBody/a:p"), "xml_missing")
      is_img <- xml_name(node) == "pic"

      if( is_table )
        add_column(pptxtable_as_tibble(node), id = id, content_type = "table cell", slide_id = slide_id)
      else if( is_par ){
        add_column(pptx_par_as_tibble(node), id = id,
                   content_type = "paragraph", slide_id = slide_id)
      } else if( is_img ){
        rel <- slide$relationship()
        images_ <- rel$get_images_path()
        img_id <- names(images_)
        images_ <- normalizePath( file.path(dirname( slide$file_name() ), images_) )
        names( images_ ) <- img_id
        add_column( embed_img_raster(node, images_), id = id, content_type = "image", slide_id = slide_id)
      } else {
        tibble( id = id, content_type = "unknown", slide_id = slide_id)
      }
    }, slide_id = i)

    content

  }, x)

}
