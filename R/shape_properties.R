# check properties helpers ----

check_set_geom <- function(x) {
  # http://www.datypic.com/sc/ooxml/t-a_ST_ShapeType.html
  geom_types <- c(
    "line", "lineInv", "triangle",
    "rtTriangle", "rect", "diamond", "parallelogram",
    "trapezoid", "nonIsoscelesTrapezoid", "pentagon", "hexagon",
    "heptagon", "octagon", "decagon", "dodecagon", "star4",
    "star5", "star6", "star7", "star8", "star10", "star12",
    "star16", "star24", "star32", "roundRect", "round1Rect",
    "round2SameRect", "round2DiagRect", "snipRoundRect",
    "snip1Rect", "snip2SameRect", "snip2DiagRect", "plaque", "ellipse",
    "teardrop", "homePlate", "chevron", "pieWedge", "pie",
    "blockArc", "donut", "noSmoking", "rightArrow",
    "leftArrow", "upArrow", "downArrow", "stripedRightArrow",
    "notchedRightArrow", "bentUpArrow", "leftRightArrow",
    "upDownArrow", "leftUpArrow", "leftRightUpArrow", "quadArrow",
    "leftArrowCallout", "rightArrowCallout", "upArrowCallout",
    "downArrowCallout", "leftRightArrowCallout",
    "upDownArrowCallout", "quadArrowCallout", "bentArrow", "uturnArrow",
    "circularArrow", "leftCircularArrow",
    "leftRightCircularArrow", "curvedRightArrow", "curvedLeftArrow",
    "curvedUpArrow", "curvedDownArrow", "swooshArrow", "cube", "can",
    "lightningBolt", "heart", "sun", "moon", "smileyFace",
    "irregularSeal1", "irregularSeal2", "foldedCorner", "bevel",
    "frame", "halfFrame", "corner", "diagStripe", "chord", "arc",
    "leftBracket", "rightBracket", "leftBrace",
    "rightBrace", "bracketPair", "bracePair", "straightConnector1",
    "bentConnector2", "bentConnector3", "bentConnector4",
    "bentConnector5", "curvedConnector2", "curvedConnector3",
    "curvedConnector4", "curvedConnector5", "callout1", "callout2",
    "callout3", "accentCallout1", "accentCallout2",
    "accentCallout3", "borderCallout1", "borderCallout2",
    "borderCallout3", "accentBorderCallout1", "accentBorderCallout2",
    "accentBorderCallout3", "wedgeRectCallout",
    "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "cloud",
    "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2",
    "leftRightRibbon", "verticalScroll", "horizontalScroll",
    "wave", "doubleWave", "plus", "flowChartProcess",
    "flowChartDecision", "flowChartInputOutput",
    "flowChartPredefinedProcess", "flowChartInternalStorage",
    "flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
    "flowChartPreparation", "flowChartManualInput",
    "flowChartManualOperation", "flowChartConnector",
    "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSummingJunction",
    "flowChartOr", "flowChartCollate", "flowChartSort",
    "flowChartExtract", "flowChartMerge", "flowChartOfflineStorage",
    "flowChartOnlineStorage", "flowChartMagneticTape",
    "flowChartMagneticDisk", "flowChartMagneticDrum",
    "flowChartDisplay", "flowChartDelay", "flowChartAlternateProcess",
    "flowChartOffpageConnector", "actionButtonBlank",
    "actionButtonHome", "actionButtonHelp", "actionButtonInformation",
    "actionButtonForwardNext", "actionButtonBackPrevious",
    "actionButtonEnd", "actionButtonBeginning",
    "actionButtonReturn", "actionButtonDocument", "actionButtonSound",
    "actionButtonMovie", "gear6", "gear9", "funnel", "mathPlus",
    "mathMinus", "mathMultiply", "mathDivide", "mathEqual",
    "mathNotEqual", "cornerTabs", "squareTabs", "plaqueTabs",
    "chartX", "chartStar", "chartPlus"
  )

  if(!x %in% geom_types) {
    stop("'", x, "' must be a valid geometry\n\n",
         "A valid geometry has to be one of ", paste0(paste0("'", geom_types, "'"), collapse = ", "), ".")
  } else {
    return(x)
  }
}

# line ----

#' @title Line properties
#'
#' @description Create a `sp_line` object that describes
#' line properties.
#'
#' @param color line color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @param lwd line width (in point) - 0 or positive integer value.
#' @param lty single character value specifying the line type.
#' Expected value is one of the following : default `'solid'`
#' or `'dot'` or `'dash'` or `'lgDash'`
#' or `'dashDot'` or `'lgDashDot'` or `'lgDashDotDot'`
#' or `'sysDash'` or `'sysDot'` or `'sysDashDot'`
#' or `'sysDashDotDot'`.
#' @param linecmpd single character value specifying the compound line type.
#' Expected value is one of the following : default `'sng'`
#' or `'dbl'` or `'tri'` or `'thinThick'`
#' or `'thickThin'`
#' @param lineend single character value specifying the line end style
#' Expected value is one of the following : default `'rnd'`
#' or `'sq'` or `'flat'`
#' @param linejoin single character value specifying the line join style
#' Expected value is one of the following : default `'round'`
#' or `'bevel'` or `'miter'`
#' @param headend a `sp_lineend` object specifying line head end style
#' @param tailend a `sp_lineend` object specifying line tail end style
#' @return a `sp_line` object
#' @examples
#' sp_line()
#' sp_line(color = "red", lwd = 2)
#' sp_line(lty = "dot", linecmpd = "dbl")
#' @family functions for defining shape properties
#' @seealso [sp_lineend]
#' @export
sp_line <- function(color = "transparent", lwd = 1, lty = "solid",
                    linecmpd = "sng", lineend = "rnd", linejoin = "round",
                    headend = sp_lineend(type = "none"), tailend = sp_lineend(type = "none")) {
  out <- list()
  out <- check_set_color(out, color)
  out <- check_set_numeric(out, lwd)
  out <- check_set_choice( obj = out, value = lty,
                           choices = c("solid", "dot", "dash",
                                       "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash",
                                       "sysDot", "sysDashDot", "sysDashDotDot") )
  out <- check_set_choice( obj = out, value = linecmpd,
                           choices = c("sng", "dbl", "tri", "thinThick", "thickThin") )
  out <- check_set_choice( obj = out, value = lineend,
                           choices = c("rnd", "sq", "flat") )
  out <- check_set_choice( obj = out, value = linejoin,
                           choices = c("round", "bevel", "miter") )
  out <- check_set_class( obj = out, value = headend, cl = "sp_lineend" )
  out <- check_set_class( obj = out, value = tailend, cl = "sp_lineend" )

  class( out ) <- c("sp_line")

  out
}

#' @param x,object `sp_line` object
#' @param ... further arguments - not used
#' @examples
#' print( sp_line (color="red", lwd = 2) )
#' @rdname sp_line
#' @export
print.sp_line = function (x, ...){
  out <- data.frame(
    color = x$color,
    lwd = x$lwd,
    lty = x$lty,
    linecmpd = x$linecmpd,
    lineend = x$lineend,
    linejoin = x$linejoin,
    headend = unclass(x$headend),
    tailend = unclass(x$tailend), stringsAsFactors = FALSE )
  print(out)
  invisible()
}

#' @rdname sp_line
#' @examples
#' obj <- sp_line (color="red", lwd = 2)
#' update( obj, linecmpd = "dbl" )
#' @export
update.sp_line <- function(object, color, lwd, lty,
                           linecmpd, lineend, linejoin,
                           headend, tailend, ...) {
  if( !missing(color) )
    object <- check_set_color(object, color)
  if( !missing(lwd) )
    object <- check_set_numeric(object, lwd)
  if( !missing(lty) )
    object <- check_set_choice( obj = object, value = lty,
                           choices = c("solid", "dot", "dash",
                                       "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash",
                                       "sysDot", "sysDashDot", "sysDashDotDot") )
  if( !missing(linecmpd) )
  object <- check_set_choice( obj = object, value = linecmpd,
                           choices = c("sng", "dbl", "tri", "thinThick", "thickThin") )
  if( !missing(lineend) )
    object <- check_set_choice( obj = object, value = lineend,
                           choices = c("rnd", "sq", "flat") )

  if( !missing(linejoin) )
    object <- check_set_choice( obj = object, value = linejoin,
                           choices = c("round", "bevel", "miter") )

  if( !missing(headend) )
    object <- check_set_class( obj = object, value = headend, cl = "sp_lineend" )
  if( !missing(tailend) )
    object <- check_set_class( obj = object, value = tailend, cl = "sp_lineend" )

  object
}

#' @title Line end properties
#'
#' @description Create a `sp_lineend` object that describes
#' line end properties.
#'
#' @param type single character value specifying the line end type.
#' Expected value is one of the following : default `'none'`
#' or `'triangle'` or `'stealth'` or `'diamond'`
#' or `'oval'` or `'arrow'`
#' @param width single character value specifying the line end width
#' Expected value is one of the following : default `'sm'`
#' or `'med'` or `'lg'`
#' @param length single character value specifying the line end length
#' Expected value is one of the following : default `'sm'`
#' or `'med'` or `'lg'`
#' @return a `sp_lineend` object
#' @examples
#' sp_lineend()
#' sp_lineend(type = "triangle")
#' sp_lineend(type = "arrow", width = "lg", length = "lg")
#' @family functions for defining shape properties
#' @seealso [sp_line]
#' @export
sp_lineend <- function(type = "none", width = "med", length = "med") {
  out <- list()
  out <- check_set_choice( obj = out, value = type,
                           choices = c("none","triangle","stealth",
                                       "diamond","oval","arrow") )
  out <- check_set_choice( obj = out, value = width,
                           choices = c("sm", "med", "lg") )
  out <- check_set_choice( obj = out, value = length,
                           choices = c("sm", "med", "lg") )

  class( out ) <- c("sp_lineend")

  out
}

#' @param x,object `sp_lineend` object
#' @param ... further arguments - not used
#' @examples
#' print(sp_lineend (type="triangle", width = "lg"))
#' @rdname sp_lineend
#' @export
print.sp_lineend = function (x, ...){
  out <- data.frame(
    type = x$type,
    width = x$width,
    length = x$length
  )
  print(out)
  invisible()
}

#' @rdname sp_lineend
#' @examples
#' obj <- sp_lineend (type="triangle", width = "lg")
#' update( obj, type = "arrow" )
#' @export
update.sp_lineend <- function(object, type, width, length,
                           ...) {
  if( !missing(type) )
    object <- check_set_choice( obj = object, value = type,
                                choices = c("none","triangle","stealth",
                                            "diamond","oval","arrow") )
  if( !missing(width) )
    object <- check_set_choice( obj = object, value = width,
                                choices = c("sm", "med", "lg") )
  if( !missing(length) )
    object <- check_set_choice( obj = object, value = length,
                                choices = c("sm", "med", "lg") )

  object
}
