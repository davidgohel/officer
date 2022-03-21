unfold_row_pml <- function(node, row_id){

  children_ <- xml_children(node)
  cell_nodes <- children_[sapply(children_, function(x) xml_name(x)=="tc" )]

  txt <- sapply(cell_nodes, xml_text)

  col_span <- sapply(cell_nodes, function(x) {
    as.integer(xml_attr(x, "gridSpan"))
  })
  h_merge <- sapply(cell_nodes, function(x) {
    as.integer(xml_attr(x, "hMerge"))
  }) %in% c(1)
  col_span[is.na(col_span)] <- 1L
  col_span[h_merge] <- 0L

  row_span <- lapply(cell_nodes, function(x) {
    row_span <- as.integer(xml_attr(x, "rowSpan"))
    v_merged <- as.integer(xml_attr(x, "vMerge"))
    data.frame(v_merged = v_merged,
               row_span = row_span,
               stringsAsFactors = FALSE
    )
  })
  row_span <- rbind.match.columns(row_span)
  row_span$row_merge <- !is.na(row_span$v_merged) | !is.na(row_span$row_span)
  row_span$first <- !is.na(row_span$row_span)
  row_span$row_span[!is.na(row_span$v_merged)] <- 0L
  row_span$row_span[!row_span$row_merge] <- 1L
  row_span <- row_span[, c("row_merge", "first", "row_span") ]


  out <- data.frame(row_id = row_id, cell_id = seq_along(cell_nodes),
                text = txt, col_span = col_span,
                row_merge = row_span$row_merge,
                first = row_span$first,
                stringsAsFactors = FALSE)
  out
}

globalVariables(c(".", "src", "width", "height"))

pptxtable_as_tibble <- function( node ){
  tbl <- xml_child(node, "/a:graphic/a:graphicData/a:tbl")
  xpath_ <- paste0( xml_path(tbl), "/a:tr")
  rows <- xml_find_all(node, xpath_)
  if( length(rows) < 1 ) return(NULL)
  row_details <- mapply(unfold_row_pml, rows, seq_along(rows), SIMPLIFY = FALSE)
  row_details <- rbind.match.columns(row_details)
  row_details <- set_row_span(row_details)
  row_details$text[row_details$col_span < 1 | row_details$row_span < 1] <- NA_character_
  row_details
}


pptx_par_as_tibble <- function(node){
  xpath_ <- paste0( xml_path( node ),
                    c("/p:txBody/a:p", "//*/p:txBody/a:p"), # standard and groupedshapes
                    collapse = "|")

  p_nodes <- xml_find_all(node, xpath_ )
  data.frame( text = xml_text(p_nodes), stringsAsFactors = FALSE )
}



embed_img_raster  <- function(node, img_src ){
  blip <- xml_child(node, "/p:blipFill/a:blip" )
  img_id <- xml_attr(blip, "embed")

  file_ <- img_src[img_id]
  stopifnot(is.character(file_), length(file_) == 1)
  data.frame(media_file = file.path( "ppt/media/", basename(file_) ),
             stringsAsFactors = FALSE )
}

#' @export
#' @title Extract media from a document object
#' @description Extract files from a \code{rpptx} object.
#' @param x an rpptx object
#' @param path media path, should be a relative path
#' @param target target file
#' @examples
#' example_pptx <- system.file(package = "officer",
#'   "doc_examples/example.pptx")
#' doc <- read_pptx(example_pptx)
#' content <- pptx_summary(doc)
#' image_row <- content[content$content_type %in% "image", ]
#' media_file <- image_row$media_file
#' png_file <- tempfile(fileext = ".png")
#' media_extract(doc, path = media_file, target = png_file)
media_extract <- function( x, path, target ){
  media <- file.path(x$package_dir, path )
  stopifnot(file.exists(media))
  file.copy(from = media, to = target)
}

#' @title PowerPoint content in a data.frame
#' @description Read content of a PowerPoint document and
#' return a dataset representing the document.
#' @param x an rpptx object
#' @examples
#' example_pptx <- system.file(package = "officer",
#'   "doc_examples/example.pptx")
#' doc <- read_pptx(example_pptx)
#' pptx_summary(doc)
#' pptx_summary(example_pptx)
#' @export
pptx_summary <- function( x ){
  if (is.character(x)) {
    x = read_pptx(x)
  }
  list_content <- list()
  for( i in seq_len( length(x) )){
    slide <- x$slide$get_slide(i)
    str = as_xpath_content_sel("p:cSld/p:spTree/")
    nodes <- xml_find_all(slide$get(), str)

    if( length(nodes) < 1 ) { # when slide is empty
      list_content[[length(list_content)+1]] <- data.frame( id = character(0), content_type = character(0), slide_id = integer(0),
                                                            stringsAsFactors = FALSE)
      next
    }

    content <- mapply(function(node, slide_id){

      is_table <- !inherits( xml_child(node, "/a:graphic/a:graphicData/a:tbl"), "xml_missing")
      is_par <- !inherits( xml_child(node, "/p:txBody/a:p"), "xml_missing")
      is_img <- xml_name(node) == "pic"
      id <- xml_attr(xml_child(node, "/p:cNvPr"), "id")

      if( is_table ){
        ppt_tab <- pptxtable_as_tibble(node)
        ppt_tab$id <- id
        ppt_tab$content_type <- "table cell"
        ppt_tab$slide_id <- slide_id
        ppt_tab
      } else if( is_par ){
        ppt_tab <- pptx_par_as_tibble(node)
        ppt_tab$id <- id
        ppt_tab$content_type <- "paragraph"
        ppt_tab$slide_id <- slide_id
        ppt_tab
      } else if( is_img ){
        rel <- slide$relationship()
        images_ <- rel$get_images_path()
        img_id <- names(images_)
        images_ <- normalizePath( file.path(dirname( slide$file_name() ), images_) )
        names( images_ ) <- img_id
        ppt_tab <- embed_img_raster(node, images_)
        ppt_tab$id <- id
        ppt_tab$content_type <- "image"
        ppt_tab$slide_id <- slide_id
        ppt_tab
      } else {
        data.frame( id = id, content_type = "unknown", slide_id = slide_id,
                    stringsAsFactors = FALSE)
      }
    }, nodes, slide_id = i, SIMPLIFY = FALSE)

    list_content[[length(list_content)+1]] <- rbind.match.columns(content)
  }

  rbind.match.columns(list_content)
}
