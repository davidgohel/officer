# tags with namespaces ----

wp_ns_yes <- "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">"
wp_ns_no <- "<w:p>"

tbl_ns_yes <- "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
tbl_ns_no <- "<w:tbl>"

wr_ns_yes <- "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">"
wr_ns_no <- "<w:r>"

ar_ns_yes <- "<a:r xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
ar_ns_no <- "<a:r>"

ap_ns_yes <- "<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
ap_ns_no <- "<a:p>"

psp_ns_yes <- "<p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
psp_ns_no <- "<p:sp>"

# utils -----
runs_to_p_wml <- function(..., add_ns = FALSE, style_id = NULL){
  runs <- list(...)
  run_str <- lapply(runs, to_wml, add_ns = FALSE)
  run_str$collapse <- ""
  run_str <- do.call(paste0, run_str)
  open_tag <- wp_ns_no
  if (add_ns) {
    open_tag <- wp_ns_yes
  }

  if( !is.null(style_id) )
    ppr <- paste0("<w:pPr><w:pStyle w:val=\"", style_id, "\"/></w:pPr>")
  else ppr <- "<w:pPr/>"
  out <- paste0(open_tag, ppr, run_str, "</w:p>")
  out
}


a_xfrm_str <- function( left = 0, top = 0, width = 3, height = 3, rot = 0){

  start_tag <- "<a:xfrm>"
  if( !is.null(rot) && !is.na(rot) && rot != 0) {
    start_tag <- sprintf("<a:xfrm rot=\"%.0f\">", -rot * 60000)
  }

  xfrm_str <- paste0(start_tag, "<a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></a:xfrm>")
  sprintf(xfrm_str,
          left * 914400, top * 914400,
          width * 914400, height * 914400)
}

p_xfrm_str <- function( left = 0, top = 0, width = 3, height = 3, rot = 0){

  if( is.null(rot) || !is.finite(rot) ) {
    rot <- 0
  }

  xfrm_str <- "<p:xfrm rot=\"%.0f\"><a:off x=\"%.0f\" y=\"%.0f\"/><a:ext cx=\"%.0f\" cy=\"%.0f\"/></p:xfrm>"
  sprintf(xfrm_str, -rot * 60000,
          left * 914400, top * 914400,
          width * 914400, height * 914400)
}


# colors ----

css_color <- function(color){
  color <- as.vector(col2rgb(color, alpha = TRUE)) / c(1, 1, 1, 255)

  if( !(color[4] > 0) ) "transparent"
  else sprintf("rgba(%.0f,%.0f,%.0f,%.2f)",
               color[1], color[2], color[3], color[4])
}
hex_color <- function(color){
  color <- as.vector(col2rgb(color, alpha = TRUE)) / c(1, 1, 1, 255)
  sprintf("%02X%02X%02X",
          color[1], color[2], color[3])
}
colalpha <- function(x){
  if("transparent" %in% x) return(0)
  alpha <- as.vector(col2rgb(x, alpha = TRUE))[4] / 255
  alpha
}

is_transparent <- function(color){
  !(colalpha(color) > 0)
}

solid_fill <- function(color) {
  sprintf(
    paste0("<a:solidFill><a:srgbClr val=\"%s\">",
           "<a:alpha val=\"%.0f\"/>",
           "</a:srgbClr></a:solidFill>"),
    hex_color(color),
    colalpha(color) * 100000 )
}

solid_fill_pml <- function(bg){
  bg_str <- ""
  if( !is.null(bg)){
    bg_str <- solid_fill(bg)
  }
  bg_str
}

# geom ----

prst_geom_pml <- function(x) {
  geom_str <- ""
  if (!is.null(x)) {
    x <- check_set_geom(x)
    tagname <- paste0("a:prstGeom")
    geom_str <- sprintf("<%s prst=\"%s\"><a:avLst/></%s>", tagname, x, tagname)
  }
  geom_str
}

# line ----

ln_pml <- function(x) {
  ln_str <- ""
  if (!is.null(x)) {
    color_ <- ""
    if(is_transparent(x$color) || x$lwd < .001) {
      color_ <- "<a:noFill/>"
    } else {
      color_ <- solid_fill(x$color)
    }

    dash_ <- dash_pml(x$lty)
    join_ <- linejoin_pml(x$linejoin)
    head_ <- lineend_pml(x$headend, "head")
    tail_ <- lineend_pml(x$tailend, "tail")

    ln_str <- sprintf(
      paste0("<a:ln w=\"%s\" cap=\"%s\" cmpd=\"%s\">",
             color_,
             dash_,
             join_,
             head_,
             tail_,
             "</a:ln>"),
      12700 * x$lwd, x$lineend, x$linecmpd)

  }
  ln_str
}

lineend_pml <- function(x, side) {
  lineend_str <- ""
  if (!is.null(x)) {
    tagname <- paste0("a:", side, "End")
    lineend_str <- sprintf("<%s type=\"%s\" w=\"%s\" len=\"%s\"/>", tagname, x$type, x$width, x$length)
  }
  lineend_str
}

dash_pml <- function(x) {
  dash_str <- ""
  if (!is.null(x)) {
    dash_str <- sprintf("<a:prstDash val=\"%s\"/>", x)
  }
  dash_str
}

linejoin_pml <- function(x) {
  linejoin_str <- ""
  if (!is.null(x)) {
    linejoin_str <- sprintf("<a:%s/>", x)
  }
  linejoin_str
}

# border ----

border_pml <- function(x, side){
  tagname <- paste0("a:ln", side)

  width_ <- sprintf("w=\"%.0f\"", x$width * 12700)

  if(is_transparent(x$color) || x$width < .001){
    color_ <- "<a:noFill/>"
  } else {
    color_ <- solid_fill(x$color)
  }

  if( !x$style %in% c("dotted", "dashed", "solid") ){
    x$style <- "solid"
  }
  if( "dotted" %in% x$style ){
    x$style <- "sysDot"
  } else if( "dashed" %in% x$style ){
    x$style <- "sysDash"
  }

  style_ <- sprintf("<a:prstDash val=\"%s\"/>", x$style)

  paste0("<", tagname,
         " ", "algn=\"ctr\" cmpd=\"sng\" cap=\"flat\"",
         " ", width_,
         ">", color_, style_,
         "</", tagname, ">"
         )
}


border_wml <- function(x, side){
  tagname <- paste0("w:", side)
  x$style[x$style %in% "solid"] <- "single"
  if( !x$style %in% c("dotted", "dashed", "single") ){
    x$style <- "single"
  }
  if( x$width < 0.0001 || is_transparent(x$color) ){
    x$style <- "none"
  }


  style_ <- sprintf("w:val=\"%s\"", x$style)
  width_ <- sprintf("w:sz=\"%.0f\"", x$width*8)
  color_ <- sprintf("w:color=\"%s\"", hex_color(x$color))

  paste0("<", tagname,
         " ", style_,
         " ", width_,
         " ", "w:space=\"0\"",
         " ", color_,
         "/>")

}
border_css <- function(x, side){

  color_ <- css_color(x$color)
  if( !(x$width > 0 ) )
    color_ <- "transparent"

  width_ <- sprintf("%.02fpt", x$width)

  if( !x$style %in% c("dotted", "dashed", "solid") ){
    x$style <- "solid"
  }
  paste0("border-", side, ": ", width_, " ", x$style, " ", color_, ";")

}

# ppr ----
ppr_pml <- function(x){
  align  <- " algn=\"r\""
  if (x$text.align == "left" ){
    align  <- " algn=\"l\"";
  } else if(x$text.align == "center" ){
    align  <- " algn=\"ctr\"";
  } else if(x$text.align == "justify" ){
    align  <- " algn=\"just\"";
  }
  leftright_padding <- sprintf(" marL=\"%.0f\" marR=\"%.0f\"", x$padding.left*12700, x$padding.right*12700)
  top_padding <- sprintf("<a:spcBef><a:spcPts val=\"%.0f\" /></a:spcBef>", x$padding.top*100)
  bottom_padding <- sprintf("<a:spcAft><a:spcPts val=\"%.0f\" /></a:spcAft>", x$padding.bottom*100)

  line_spacing <- sprintf("<a:lnSpc><a:spcPct val=\"%.0f\"/></a:lnSpc>", x$line_spacing*100000)

  paste0("<a:pPr", align, leftright_padding, ">",
         line_spacing,
         top_padding, bottom_padding, "<a:buNone/>",
         "</a:pPr>")
}

ppr_css <- function(x){

  text.align  <- sprintf("text-align:%s;", x$text.align)
  borders <- paste0(
    border_css(x$border.bottom, "bottom"),
    border_css(x$border.top, "top"),
    border_css(x$border.left, "left"),
    border_css(x$border.right, "right") )

  paddings <- sprintf("padding-top:%.0fpt;padding-bottom:%.0fpt;padding-left:%.0fpt;padding-right:%.0fpt;",
                      x$padding.top, x$padding.bottom, x$padding.left, x$padding.right)
  ls <- formatC(x$line_spacing, format = "f", digits = 2, decimal.mark = ".", drop0trailing = TRUE )
  line_spacing <- sprintf("line-height: %s;", ls )

  shading.color <- sprintf("background-color:%s;", css_color(x$shading.color))

  paste0("margin:0pt;", text.align, borders, paddings, line_spacing,
         shading.color)
}

ppr_wml <- function(x){

  if("justify" %in% x$text.align ){
    x$text.align  <- "both";
  }
  pstyle <- ""
  if(!is.null(x$word_style)) {
    word_style_id <- gsub("[^a-zA-Z0-9]", "", x$word_style)
    pstyle <- sprintf("<w:pStyle w:val=\"%s\"/>", word_style_id)
  }
  text_align_ <- sprintf("<w:jc w:val=\"%s\"/>", x$text.align)
  keep_with_next <- ""
  if(x$keep_with_next){
    keep_with_next <- "<w:keepNext/>"
  }
  borders_ <- paste0("<w:pBdr>",
    border_wml(x$border.bottom, "bottom"),
    border_wml(x$border.top, "top"),
    border_wml(x$border.left, "left"),
    border_wml(x$border.right, "right"), "</w:pBdr>" )

  leftright_padding <- sprintf("<w:ind w:left=\"%.0f\" w:right=\"%.0f\" w:firstLine=\"0\" w:firstLineChars=\"0\"/>",
                               x$padding.left*20, x$padding.right*20)
  topbot_spacing <- sprintf("<w:spacing w:after=\"%.0f\" w:before=\"%.0f\" w:line=\"%.0f\"/>",
                            x$padding.bottom*20, x$padding.top*20, x$line_spacing*240)
  shading_ <- ""
  if(!is_transparent(x$shading.color)){
    shading_ <- sprintf(
      "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>",
      hex_color(x$shading.color))
  }

  paste0("<w:pPr>",
         pstyle,
         text_align_,
         keep_with_next,
         borders_,
         topbot_spacing,
         leftright_padding,
         shading_,
         "</w:pPr>")

}


# rpr ----

rpr_pml <- function(x){

  if(is_transparent(x$color) ) return("")

  out  <- "<a:rPr cap=\"none\""

  if(!is.na(x$font.size)){
    if( x$font.size > 0 ){
      out <- paste0(out, sprintf(" sz=\"%.0f\"", x$font.size * 100) )
    }
  }

  if(!is.na(x$italic)){
    if(x$italic){
      out <- paste0(out, " i=\"1\"")
    } else {
      out <- paste0(out, " i=\"0\"")
    }
  }

  if(!is.na(x$bold)){
    if(x$bold){
      out <- paste0(out, " b=\"1\"")
    } else {
      out <- paste0(out, " b=\"0\"")
    }
  }

  if(!is.na(x$underlined)){
    if(x$underlined){
      out <- paste0(out, " u=\"sng\"")
    } else {
      out <- paste0(out, " u=\"none\"")
    }
  }

  if( x$vertical.align == "superscript"){
    out <- paste0(out, " baseline=\"40000\"")
  } else if( x$vertical.align == "subscript"){
    out <- paste0(out, " baseline=\"-40000\"")
  }
  out <- paste0(out, ">")

  if(!is.na(x$color)){
    out <- paste0(out, solid_fill(x$color))

    if(!is_transparent(x$shading.color) ){
      shad <- sprintf(
        paste0("<a:highlight><a:srgbClr val=\"%s\">",
               "<a:alpha val=\"%.0f\"/>",
               "</a:srgbClr></a:highlight>"),
        hex_color(x$shading.color),
        colalpha(x$shading.color) * 100000 )
      out <- paste0(out, shad)
    }
  }

  out <- paste0(
    out,
    if(!is.na(x$font.family)) sprintf("<a:latin typeface=\"%s\"/>", x$font.family),
    if(!is.na(x$cs.family)) sprintf("<a:cs typeface=\"%s\"/>", x$cs.family),
    if(!is.na(x$eastasia.family)) sprintf("<a:ea typeface=\"%s\"/>", x$eastasia.family),
    if(!is.na(x$hansi.family)) sprintf("<a:sym typeface=\"%s\"/>", x$hansi.family)
  )

  out <- paste0(out, "</a:rPr>")
  out
}

rpr_wml <- function(x){

  out <- paste0(
    "<w:rPr><w:rFonts",
    if (!is.na(x$font.family)) paste0(" w:ascii=\"", x$font.family, "\""),
    if (!is.na(x$hansi.family)) paste0(" w:hAnsi=\"", x$hansi.family, "\""),
    if (!is.na(x$eastasia.family)) paste0(" w:eastAsia=\"", x$eastasia.family, "\""),
    if (!is.na(x$cs.family)) paste0(" w:cs=\"", x$cs.family, "\""),
    "/>"
  )

  if (!is.na(x$italic)) {
    if (x$italic) {
      out <- paste0(out, "<w:i w:val=\"true\"/>")
    } else {
      out <- paste0(out, "<w:i w:val=\"false\"/>")
    }
  }

  if (!is.na(x$bold)) {
    if (x$bold) {
      out <- paste0(out, "<w:b w:val=\"true\"/>")
    } else {
      out <- paste0(out, "<w:b w:val=\"false\"/>")
    }
  }
  if (!is.na(x$underlined)) {
    if (x$underlined) {
      out <- paste0(out, "<w:u w:val=\"single\"/>")
    } else {
      out <- paste0(out, "<w:u w:val=\"none\"/>")
    }
  }

  if( x$vertical.align == "superscript"){
    out <- paste0(out, "<w:vertAlign w:val=\"superscript\"/>")
  } else if( x$vertical.align == "subscript"){
    out <- paste0(out, "<w:vertAlign w:val=\"subscript\"/>")
  }

  if(!is.na(x$font.size)){
    out <- paste0(
      out,
      sprintf("<w:sz w:val=\"%.0f\"/><w:szCs w:val=\"%.0f\"/>",
              x$font.size * 2, x$font.size * 2)
    )
  }

  if(!is.na(x$color)){
    out <- paste0(
      out,
      sprintf("<w:color w:val=\"%s\"/>", hex_color(x$color))
    )
  }

  if(!is.na(x$shading.color)){
    if(!is_transparent(x$shading.color) ){
      out <- paste0(
        out,
        sprintf("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>", hex_color(x$shading.color))
      )
    }
    if(colalpha(x$color) < 1){
      out <- paste0(
        out,
        sprintf(
          paste0("<w14:textFill><w14:solidFill><w14:srgbClr val=\"%s\">",
                 "<w14:alpha val=\"%.0f\"/>",
                 "</w14:srgbClr></w14:solidFill></w14:textFill>"),
          hex_color(x$color),
          colalpha(x$color) * 100000 )
      )
    }
  }

  out <- paste0(out, "</w:rPr>")
  out
}

rpr_css <- function(x){

  out <- ""

  if(!is.na(x$font.family)) out <- paste0(out, sprintf("font-family:'%s';", x$font.family))
  if(!is.na(x$color)) out <- paste0(out, sprintf("color:%s;", css_color(x$color)))
  if(!is.na(x$font.size)) out <- paste0(out, sprintf("font-size:%0.1fpt;", x$font.size))

  if(!is.na(x$italic)){
    if(x$italic){
      out <- paste0(out, "font-style:italic;")
    } else {
      out <- paste0(out, "font-style:normal;")
    }
  }

  if(!is.na(x$bold)){
    if(x$bold){
      out <- paste0(out, "font-weight:bold;")
    } else {
      out <- paste0(out, "font-weight:normal;")
    }
  }

  if(!is.na(x$bold)){
    if(x$underlined){
      out <- paste0(out, "text-decoration:underline;")
    } else {
      out <- paste0(out, "text-decoration:none;")
    }
  }

  if(!is.na(x$shading.color)){
    if(!is_transparent(x$shading.color) ){
      out <- paste0(
        out,
        sprintf("background-color:%s;", css_color(x$shading.color))
      )
    } else out <- paste0(out, "background-color:transparent;")
  }

  if( x$vertical.align == "superscript"){
    out <- paste0(out, "vertical-align: super;")
  } else if( x$vertical.align == "subscript"){
    out <- paste0(out, "vertical-align: sub;")
  }


  out
}
# tcpr ----

tcpr_pml <- function(x){

  text.direction <-
    if(x$text.direction %in% "btlr")
      " vert=\"vert270\""
    else if(x$text.direction %in% "tbrl")
      " vert=\"vert\""
    else ""

  vertical.align <-
    if(x$vertical.align %in% "center"){
      " anchor=\"ctr\""
    } else if( x$vertical.align %in% "top"){
      " anchor=\"t\""
    } else " anchor=\"b\""

  margins <- sprintf(" marB=\"%.0f\" marT=\"%.0f\" marR=\"%.0f\" marL=\"%.0f\"",
                     x$margin.bottom * 12700, x$margin.top * 12700,
                     x$margin.right * 12700, x$margin.left * 12700)

  background.color <- paste0(
    sprintf("<a:solidFill><a:srgbClr val=\"%s\">", hex_color(x$background.color) ),
    sprintf("<a:alpha val=\"%.0f\"/>", colalpha(x$background.color)*100000 ),
    "</a:srgbClr></a:solidFill>" )

  bb <- border_pml(x$border.bottom, side = "B")
  bt <- border_pml(x$border.top, side = "T")
  bl <- border_pml(x$border.left, side = "L")
  br <- border_pml(x$border.right, side = "R")

  pml_attrs <- paste0(text.direction, vertical.align, margins)
  paste0("<a:tcPr", pml_attrs, ">", bl, br, bt, bb,
         background.color, "</a:tcPr>" )
}

tcpr_wml <- function(x){

  background.color <- sprintf("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>", hex_color(x$background.color) )
  vertical.align <- ifelse( x$vertical.align %in% c("center", "top"), sprintf("<w:vAlign w:val=\"%s\"/>", x$vertical.align), "<w:vAlign w:val=\"bottom\"/>" )
  text.direction <- ifelse(
    x$text.direction %in% "btlr", "<w:textDirection w:val=\"btLr\"/>",
    ifelse(x$text.direction %in% "tbrl", "<w:textDirection w:val=\"tbRl\"/>", "") )

  bb <- border_wml(x$border.bottom, side = "bottom")
  bt <- border_wml(x$border.top, side = "top")
  bl <- border_wml(x$border.left, side = "left")
  br <- border_wml(x$border.right, side = "right")

  margin.bottom <- sprintf("<w:bottom w:w=\"%.0f\" w:type=\"dxa\"/>", x$margin.bottom * 20 )
  margin.top <- sprintf("<w:top w:w=\"%.0f\" w:type=\"dxa\"/>", x$margin.top * 20 )
  margin.left <- sprintf("<w:left w:w=\"%.0f\" w:type=\"dxa\"/>", x$margin.left * 20 )
  margin.right <- sprintf("<w:right w:w=\"%.0f\" w:type=\"dxa\"/>", x$margin.right * 20 )

  rowspan <- ""
  if (x$rowspan>1) {
    rowspan <- paste0("<w:gridSpan w:val=\"", x$rowspan,"\"/>")
  }
  colspan <- ""
  if (x$colspan>1) {
    colspan <- "<w:vMerge w:val=\"restart\"/>"
  } else if (x$colspan<1) {
    colspan <- "<w:vMerge/>"
  }

  paste0("<w:tcPr>", rowspan, colspan,
         "<w:tcBorders>", bb, bt, bl, br, "</w:tcBorders>",
                         background.color,
                         "<w:tcMar>", margin.top, margin.bottom, margin.left, margin.right, "</w:tcMar>",
                         text.direction, vertical.align, "</w:tcPr>" )

}

css_px <- function(x, format = "%.0fpx"){
  ifelse( is.na(x), "inherit",
          ifelse( x < 0.001, "0", sprintf(format, x)) )
}

tcpr_css <- function(x){

  background.color <- ifelse( colalpha(x$background.color) > 0,
                              sprintf("background-clip: padding-box;background-color:%s;", css_color(x$background.color) ),
                              "background-color:transparent;")

  width <- ifelse( is.null(x$width) || is.na(x$width), "", sprintf("width:%s;", css_px(x$width * 72) ) )
  height <- ifelse( is.null(x$height) || is.na(x$height), "", sprintf("height:%s;", css_px(x$height * 72 ) ) )
  vertical.align <- ifelse(
    x$vertical.align %in% "center", "vertical-align:middle;",
    ifelse(x$vertical.align %in% "top", "vertical-align:top;", "vertical-align:bottom;") )

  bb <- border_css(x$border.bottom, side = "bottom")
  bt <- border_css(x$border.top, side = "top")
  bl <- border_css(x$border.left, side = "left")
  br <- border_css(x$border.right, side = "right")

  margin.bottom <- sprintf("margin-bottom:%s;", sprintf("%.0fpt", x$margin.bottom) )
  margin.top <- sprintf("margin-top:%s;", sprintf("%.0fpt", x$margin.top) )
  margin.left <- sprintf("margin-left:%s;", sprintf("%.0fpt", x$margin.left) )
  margin.right <- sprintf("margin-right:%s;", sprintf("%.0fpt", x$margin.right) )

  paste0(width, height, background.color, vertical.align, bb, bt, bl, br,
                         margin.bottom, margin.top, margin.left, margin.right)

}

