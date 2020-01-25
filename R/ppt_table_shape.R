table_pptx <- function(x, style_id, col_width, row_height,
                      first_row = TRUE, first_column = FALSE,
                      last_row = FALSE, last_column = FALSE, header = TRUE ){

  str <- paste0("<a:tbl>",
         sprintf("<a:tblPr firstRow=\"%.0f\" lastRow=\"%.0f\" firstColumn=\"%.0f\" lastColumn=\"%.0f\"",
                 first_row, last_row, first_column, last_column),
         ">",
         sprintf("<a:tableStyleId>%s</a:tableStyleId>", style_id),
         "</a:tblPr>",
         "<a:tblGrid>",
         paste0(sprintf("<a:gridCol w=\"%.0f\"/>", rep(col_width, length(x))), collapse = ""),
         "</a:tblGrid>")

  as_tc <- function(x) {
    paste0("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>",
           htmlEscapeCopy(enc2utf8(x)),
           "</a:t></a:r></a:p></a:txBody></a:tc>"
    )
  }
  header_str <- character(length = 0L)
  if( header ){
    header_str  <- paste0(
      sprintf("<a:tr h=\"%.0f\">", row_height),
      paste0(as_tc(colnames(x)), collapse = ""),
      "</a:tr>"
    )
  }

  z <- lapply(x, as_tc)
  z <- do.call(paste0, z)
  z <- paste0(sprintf("<a:tr h=\"%.0f\">", row_height), z, "</a:tr>", collapse = "")

  paste0(str, header_str, z, "</a:tbl>")
}

table_shape <- function(x, value, left, top, width, height,
                        first_row = TRUE, first_column = FALSE,
                        last_row = FALSE, last_column = FALSE, header = TRUE ){
  stopifnot(is.data.frame(value))

  slide <- x$slide$get_slide(x$cursor)

  def_height <- height / (nrow(value) + 1)
  def_width <- width / (ncol(value))

  value <- characterise_df(value)

  style_id <- x$table_styles$def[1]
  xml_elt <- table_pptx(value, style_id = style_id,
                       col_width = as.integer(def_width),
                       row_height = as.integer(def_height),
                       first_row = first_row, last_row = last_row,
                       first_column = first_column, last_column = last_column, header = header )

  xml_elt <- paste0(
    "<p:graphicFrame ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
    "<p:nvGraphicFramePr>",
    "<p:cNvPr id=\"\" name=\"\"/>",
    "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"true\"/></p:cNvGraphicFramePr>",
    "<p:nvPr/>",
    "</p:nvGraphicFramePr>",
    "<p:xfrm rot=\"0\">",
    sprintf("<a:off x=\"%.0f\" y=\"%.0f\"/>", left, top),
    sprintf("<a:ext cx=\"%.0f\" cy=\"%.0f\"/>", width, height),
    "</p:xfrm>",
    "<a:graphic>",
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">",
    xml_elt,
    "</a:graphicData>",
    "</a:graphic>",
    "</p:graphicFrame>"
  )
  xml_elt
}

