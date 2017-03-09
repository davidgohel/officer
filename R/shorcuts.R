#' @title shortcuts for formatting properties
#'
#' @description
#' Shortcuts for \code{fp_text}, \code{fp_par},
#' \code{fp_cell} and \code{fp_border}.
#' @name shortcuts
#' @examples
#' shortcuts$fp_bold()
#' shortcuts$fp_italic()
#' shortcuts$b_null()
NULL

slip_in_plotref <- function( x, style = NULL, depth ){
  x <- slip_in_text(x, str = ": ", style = style, pos = "before")
  x <- slip_in_seqfield(x, str = "SEQ graph \u005C* Arabic \u005Cs 1 \u005C* MERGEFORMAT", style = style, pos = "before")
  x <- slip_in_text(x, str = " - Fig. ", style = style, pos = "before")
  x <- slip_in_seqfield(x, str = sprintf("STYLEREF %.0f \u005Cs", depth), style = style, pos = "before")
  x
}

slip_in_tableref <- function( x, style = NULL, depth ){
  x <- slip_in_text(x, str = ": ", style = style, pos = "before")
  x <- slip_in_seqfield(x, str = "SEQ table \u005C* Arabic \u005Cs 1 \u005C* MERGEFORMAT", style = style, pos = "before")
  x <- slip_in_text(x, str = " - Table ", style = style, pos = "before")
  x <- slip_in_seqfield(x, str = sprintf("STYLEREF %.0f \u005Cs", depth), style = style, pos = "before")
  x
}


#' @rdname shortcuts
#' @format NULL
#' @docType NULL
#' @keywords NULL
#' @export
shortcuts <- list(
  fp_bold = function(...) {fp_text( bold = TRUE, ... )},
  fp_italic = function( ... ) {fp_text( italic = TRUE, ... )},
  b_null = function( ... )  {fp_border( width = 0, ... )},
  slip_in_plotref = slip_in_plotref,
  slip_in_tableref = slip_in_tableref
)


