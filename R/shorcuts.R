#' @title shortcuts for formatting properties
#'
#' @description
#' Shortcuts for [fp_text()], [fp_par()], [fp_cell()] and [fp_border()].
#' @name shortcuts
#' @examples
#' shortcuts$fp_bold()
#' shortcuts$fp_italic()
#' shortcuts$b_null()
NULL

#' @rdname shortcuts
#' @format NULL
#' @docType NULL
#' @keywords NULL
#' @export
shortcuts <- list(
  fp_bold = function(...) {fp_text( bold = TRUE, ... )},
  fp_italic = function( ... ) {fp_text( italic = TRUE, ... )},
  b_null = function( ... )  {fp_border( width = 0, ... )}
)


