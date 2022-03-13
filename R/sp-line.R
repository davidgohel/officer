# check properties helpers ----

check_set_class <- function( obj, value, cl){
  varname <- as.character(substitute(value))
  cl_str <- sprintf(" must be a %s object.", cl)
  if( !inherits( value, cl ) )
    stop(varname, cl_str, call. = FALSE)
  else obj[[varname]] <- value
  obj
}

# line ----

#' @title Line properties
#'
#' @description Create a \code{sp_line} object that describes
#' line properties.
#'
#' @param color line color - a single character value specifying
#' a valid color (e.g. "#000000" or "black").
#' @param lwd line width (in point) - 0 or positive integer value.
#' @param lty single character value specifying the line type.
#' Expected value is one of the following : default \code{'solid'}
#' or \code{'dot'} or \code{'dash'} or \code{'lgDash'}
#' or \code{'dashDot'} or \code{'lgDashDot'} or \code{'lgDashDotDot'}
#' or \code{'sysDash'} or \code{'sysDot'} or \code{'sysDashDot'}
#' or \code{'sysDashDotDot'}.
#' @param linecmpd single character value specifying the compound line type.
#' Expected value is one of the following : default \code{'sng'}
#' or \code{'dbl'} or \code{'tri'} or \code{'thinThick'}
#' or \code{'thickThin'}
#' @param lineend single character value specifying the line end style
#' Expected value is one of the following : default \code{'rnd'}
#' or \code{'sq'} or \code{'flat'}
#' @param linejoin single character value specifying the line join style
#' Expected value is one of the following : default \code{'round'}
#' or \code{'bevel'} or \code{'miter'}
#' @param headend a \code{sp_lineend} object specifying line head end style
#' @param tailend a \code{sp_lineend} object specifying line tail end style
#' @return a \code{sp_line} object
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

#' @param x,object \code{sp_line} object
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
#' @description Create a \code{sp_lineend} object that describes
#' line end properties.
#'
#' @param type single character value specifying the line end type.
#' Expected value is one of the following : default \code{'none'}
#' or \code{'triangle'} or \code{'stealth'} or \code{'diamond'}
#' or \code{'oval'} or \code{'arrow'}
#' @param width single character value specifying the line end width
#' Expected value is one of the following : default \code{'sm'}
#' or \code{'med'} or \code{'lg'}
#' @param length single character value specifying the line end length
#' Expected value is one of the following : default \code{'sm'}
#' or \code{'med'} or \code{'lg'}
#' @return a \code{sp_lineend} object
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


