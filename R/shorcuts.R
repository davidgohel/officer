#' @title shortcuts for formatting properties
#'
#' @description
#' Shortcuts for \code{fp_text}, \code{fp_par},
#' \code{fp_cell} and \code{fp_border}.
#' @param ... further arguments passed to original functions.
#' @rdname shortcut_properties
#' @aliases shortcut_properties


#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_bold()
fp_bold = function(...) {fp_text( bold = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_italic()
fp_italic = function( ... ) {fp_text( italic = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' b_null()
b_null = function( ... )  {fp_border( width = 0, ... )}

