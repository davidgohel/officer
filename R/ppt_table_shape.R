table_shape <- function(x, value, left, top, width, height,
                        first_row = TRUE, first_column = FALSE,
                        last_row = FALSE, last_column = FALSE, header = TRUE ){
  stopifnot(is.data.frame(value))

  slide <- x$slide$get_slide(x$cursor)

  def_height <- height*914400 / (nrow(value) + 1)
  def_width <- width*914400 / (ncol(value))

  value <- characterise_df(value)

  style_id <- x$table_styles$def[1]
  xml_elt <- pml_table(value, style_id = style_id,
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
    sprintf( "<a:off x=\"%.0f\" y=\"%.0f\"/>", left*914400, top*914400 ),
    sprintf( "<a:ext cx=\"%.0f\" cy=\"%.0f\"/>", width*914400, height*914400 ),
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

