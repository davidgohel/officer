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
#' t_bold()
fp_bold = function(...) {fp_text( bold = TRUE, ... )}




#' @rdname shortcut_properties
#' @export
#' @examples
#' t_italic()
fp_italic = function( ... ) {fp_text( italic = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' t_bi()
fp_bi = function( ... ) {fp_text( bold = TRUE, italic = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_right()
fp_right = function( ... ) {fp_par( text.align = "right", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_left()
fp_left = function( ... ) {fp_par( text.align = "left", ... )}

#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_center()
fp_center = function( ... ) {fp_par( text.align = "center", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' fp_justify()
fp_justify = function( ... ) {fp_par( text.align = "justify", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' b_dot()
b_dot = function( ... )  {fp_border( style="dotted", ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' b_dash()
b_dash = function( ... )  {fp_border( style="dashed", ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' b_null()
b_null = function( ... )  {fp_border( width = 0, ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_null()
c_b_null = function( ... )  {
  fp_cell( border = b_null(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_all()
c_b_all = function( ... )  {
  fp_cell( border = fp_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_b()
c_b_b = function( ... )  {
  fp_cell( border = b_null(), border.bottom = fp_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_t()
c_b_t = function( ... )  {
  fp_cell( border = b_null(), border.top = fp_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_tb()
c_b_tb = function( ... ) {
  fp_cell( border = b_null(),
           border.bottom = fp_border(),
           border.top = fp_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_tb()
c_b_lr = function( ... ) {
  fp_cell( border = b_null(),
           border.left = fp_border(),
           border.right = fp_border(), ... )
}
