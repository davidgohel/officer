#' Layout selection helper
#'
#' Select a layout by name or index. The master name is inferred and only required
#' for disambiguation in case the layout name is not unique across masters.
#'
#' @param x An `rpptx` object.
#' @param layout Layout name or index. Index refers to the row index of the [layout_summary()]
#' output.
#' @param master Name of master. Only required if layout name is not unique across masters.
#' @return A `<layout_info>` object, i.e. a list with the entries `index`, `layout_name`,
#' `layout_file`, `master_name`, `master_file`, and `slide_layout`.
#' @keywords internal
get_layout <- function(x, layout, master = NULL) {
  stop_if_not_rpptx(x, "x")
  if (!(is.numeric(layout) || is.character(layout))) {
    cli::cli_abort(
      c("{.arg layout} must be {.cls numeric} or {.cls character}",
        "x" = "Got class {.cls {class(layout)[1]}} instead"
      )
    )
  }
  if (length(layout) != 1) {
    cli::cli_abort(
      c("{.arg layout} is not length 1",
        "x" = "{.arg layout} must be {.emph one} layout name or index."
      )
    )
  }
  df <- x$slideLayouts$get_metadata()
  names(df)[2:3] <- c("layout_name", "layout_file") # consistent naming
  n_layouts <- nrow(df)
  df$index <- seq_len(n_layouts)

  if (n_layouts == 0) {
    cli::cli_alert_danger("No layouts available.")
    return(NULL)
  }

  if (is.numeric(layout)) {
    res <- get_row_by_index(df, layout)
  } else {
    res <- get_row_by_name(df, layout, master)
  }
  l <- as.list(res)
  slide_layout <- x$slideLayouts$collection_get(l$layout_file)
  l <- c(l, slide_layout = slide_layout)
  l <- l[c("index", "layout_name", "layout_file", "master_name", "master_file", "slide_layout")] # nice order
  class(l) <- c("layout_info", "list")
  l
}


#' @export
print.layout_info <- function(x, ...) {
  cli::cli_h3("{.cls layout_info} object")
  str(utils::head(x, -1), give.attr = FALSE, no.list = TRUE)
  cat(" $ slide_layout: 'R6' <slide_layout>")
}


get_row_by_index <- function(df, layout) {
  index <- layout
  if (!index %in% df$index) {
    cli::cli_abort(
      c("Layout index out of bounds.",
        "x" = "Index must be between {.val {1}} and {.val {nrow(df)}}.",
        "i" = cli::col_grey("See row indexes in {.fn layout_summary}")
      ),
      call = NULL
    )
  }
  df[index, ]
}


# select layout by name
get_row_by_name <- function(df, layout, master) {
  if (!is.null(master)) {
    masters <- unique(df$master_name)
    if (!master %in% masters) {
      cli::cli_abort(c(
        "master {.val {master}} does not exist.",
        "i" = "See {.fn layout_summary} for available masters."
      ), call = NULL)
    }
    df <- df[df$master_name == master, ]
  }

  df <- df[df$layout_name == layout, ]
  if (nrow(df) == 0) {
    msg <- ifelse(is.null(master),
      "Layout {.val {layout}} does not exist",
      "Layout {.val {layout}} does not exist in master {.val {master}}"
    )
    cli::cli_abort(c(msg, "i" = "See {.fn layout_summary} for available layouts."), call = NULL)
    return(NULL)
  }
  if (nrow(df) > 1) {
    cli::cli_abort(c(
      "Layout exists in more than one master",
      "x" = "Please specify the master name in arg {.arg master}"
    ), call = NULL)
  }
  df
}


# get <layout_info> object for slide layout
get_slide_layout <- function(x, slide_idx) {
  stop_if_not_rpptx(x)
  if (length(x) == 0) {
    cli::cli_abort(
      c("Presentation does not have any slides yet",
        "x" = "Can only get the layout for an existing slides",
        "i" = "You can add a slide using {.fn add_slide}")
      , call = NULL)
  }
  ensure_slide_index_exists(x, slide_idx)
  df <- x$slide$get_xfrm()[[slide_idx]]
  layout <- unique(df$name)
  master <- unique(df$master_name)
  get_layout(x, layout, master)
}


# get <layout_info> object for layout of current slide
get_layout_for_current_slide <- function(x) {
  get_slide_layout(x, x$cursor)
}
