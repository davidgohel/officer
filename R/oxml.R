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
  if("transparent" %in% color) return(0)
  alpha <- as.vector(col2rgb(color, alpha = TRUE))[4] / 255
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
# std::string border::a_tag(std::string side)
# {
#   color_spec col_(this->red, this->green, this->blue, this->alpha);
#
#   /*if( col_.is_transparent() || width < 1 ) {
#     return "";
#   }*/
#
#     std::stringstream os;
#
#   os << "<a:ln" << side << " algn=\"ctr\" cmpd=\"sng\" cap=\"flat\" ";
#   os << "w=\"";
#   os << std::setprecision(0) << std::fixed << width * 12700;
#   os << "\">";
#
#   if( col_.is_transparent() || width < .001 ) {
#     os << "<a:noFill/>";
#   } else os << col_.solid_fill();
#
#   os << "<a:prstDash val=";
#   if( type == "solid")
#     os << "\"solid\"/>";
#   else if( type == "dotted")
#     os << "\"sysDot\"/>";
#   else if( type == "dashed")
#     os << "\"sysDash\"/>";
#
#   os << "</a:ln" << side << ">";
#   return os.str();
# }




border_wml <- function(x, side){
  tagname <- paste0("w:", side)
  if( !x$style %in% c("dotted", "dashed", "solid") ){
    x$style <- "solid"
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
  paste0("border-", side, ": ", width_, " ", x$style, " ", color_)

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

  paste0("<a:pPr", align, leftright_padding, ">",
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

  shading.color <- sprintf("background-color:%s;", css_color(x$shading.color))

  paste0("margin:0pt;", text.align, borders, paddings,
         shading.color)
}

ppr_wml <- function(x){

  if("justify" %in% x$text.align ){
    x$text.align  <- "both";
  }
  text_align_ <- sprintf("<w:jc w:val=\"%s\"/>", x$text.align)
  borders_ <- paste0(
    border_wml(x$border.bottom, "bottom"),
    border_wml(x$border.top, "top"),
    border_wml(x$border.left, "left"),
    border_wml(x$border.right, "right") )

  leftright_padding <- sprintf("<w:ind w:firstLine=\"0\" w:left=\"%.0f\" w:right=\"%.0f\"/>",
                               x$padding.left*20, x$padding.right*20)
  topbot_spacing <- sprintf("<w:spacing w:after=\"%.0f\" w:before=\"%.0f\"/>",
                            x$padding.bottom*20, x$padding.top*20)
  shading_ <- ""
  if(!is_transparent(x$shading.color)){
    shading_ <- sprintf(
      "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>",
      hex_color(x$shading.color))
  }

  paste0("<w:pPr>",
         text_align_,
         borders_,
         topbot_spacing,
         leftright_padding,
         shading_,
         "</w:pPr>")

}

