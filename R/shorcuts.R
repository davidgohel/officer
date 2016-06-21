#' @title shortcuts for formatting properties
#'
#' @description
#' Shortcuts for \code{pr_text}, \code{pr_par},
#' \code{pr_cell} and \code{pr_border}.
#' @param ... further arguments passed to original functions.
#' @rdname shortcut_properties
#' @aliases shortcut_properties


#' @rdname shortcut_properties
#' @export
#' @examples
#' t_b()
t_b = function(...) {pr_text( bold = TRUE, ... )}




#' @rdname shortcut_properties
#' @export
#' @examples
#' t_i()
t_i = function( ... ) {pr_text( italic = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' t_bi()
t_bi = function( ... ) {pr_text( bold = TRUE, italic = TRUE, ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' p_r()
p_r = function( ... ) {pr_par( text.align = "right", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' p_l()
p_l = function( ... ) {pr_par( text.align = "left", ... )}

#' @rdname shortcut_properties
#' @export
#' @examples
#' p_c()
p_c = function( ... ) {pr_par( text.align = "center", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' p_j()
p_j = function( ... ) {pr_par( text.align = "justify", ... )}


#' @rdname shortcut_properties
#' @export
#' @examples
#' b_dot()
b_dot = function( ... )  {pr_border( style="dotted", ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' b_dash()
b_dash = function( ... )  {pr_border( style="dashed", ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' b_null()
b_null = function( ... )  {pr_border( width = 0, ... )}



#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_null()
c_b_null = function( ... )  {
  pr_cell( border = b_null(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_all()
c_b_all = function( ... )  {
  pr_cell( border = pr_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_b()
c_b_b = function( ... )  {
  pr_cell( border = b_null(), border.bottom = pr_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_t()
c_b_t = function( ... )  {
  pr_cell( border = b_null(), border.top = pr_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_tb()
c_b_tb = function( ... ) {
  pr_cell( border = b_null(),
           border.bottom = pr_border(),
           border.top = pr_border(), ... )
}

#' @rdname shortcut_properties
#' @export
#' @examples
#' c_b_tb()
c_b_lr = function( ... ) {
  pr_cell( border = b_null(),
           border.left = pr_border(),
           border.right = pr_border(), ... )
}
