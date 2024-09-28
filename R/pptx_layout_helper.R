#' Layout selection helper
#'
#' Select a layout by name or index. The master name is inferred and only required
#' for disambiguation in case the layout name is not unique across masters.
#'
#' @param x An `rpptx` object.
#' @param layout Layout name or index, see [layout_summary()].
#' @param master Name of master. Only required if layout name is not unique across masters.
#' @return A `<layout_info>` object, i.e. a list with the entries `index`, `layout_name`,
#' `layout_file`, `master_name`, `master_file`, and `slide_layout`.
#' @keywords internal
#' @examples
#' x <- read_pptx()
#' get_layout(x, "Title Slide")
#'
get_layout <- function(x, layout, master = NULL) {
  if (!(is.numeric(layout) || is.character(layout))) {
    cli::cli_abort(
      c("{.arg layout} must be {.cls numeric} or {.cls character}",
        "x" = "Got class {.cls {class(layout)[1]}} instead"
      )
    )
  }
  df <- x$slideLayouts$get_metadata()
  names(df)[2:3] <- c("layout_name", "layout_file") # consistent naming
  n_layouts <- nrow(df)

  if (n_layouts == 0) {
    cli::cli_alert_danger("No layouts available.")
    return(NULL)
  }

  if (is.numeric(layout)) {
    index <- layout
    if (!index %in% seq_len(n_layouts)) {
      cli::cli_abort(
        c("Layout with index {.val {index}} does not exist.",
          "x" = "Must be between {.val {1}} and {.val {n_layouts}}.",
          "i" = cli::col_grey("See rownames in {.fn layout_summary}")
        ),
        call = NULL
      )
    }
  } else {
    layout_exists(x, layout, must_exist = TRUE)
    layout_is_unique(x, layout, require_unique = TRUE)
    index <- which(df$layout_name == layout)
  }

  l <- df[index, ] |> as.list()
  slide_layout <- x$slideLayouts$collection_get(l$layout_file)
  l <- c(index = index, l, slide_layout = slide_layout)
  l <- l[c("index", "layout_name", "layout_file", "master_name", "master_file", "slide_layout")] # nice order
  class(l) <- c("layout_info", "list")
  l
}


#' @export
print.layout_info <- function(x, ...) {
  cli::cli_h3("{.cls layout_info} object")
  str(head(x, -1), give.attr = FALSE)
  cat(" $ slide_layout: 'R6' <slide_layout>")
}


get_layout_names <- function(x) {
  layout_summary(x)$layout # layout_summary as used in print.rpptx
}


# find similar name
find_similar_names <- function(name, names, n = 3) {
  if (!requireNamespace("stringdist", quietly = TRUE)) {
    return(NULL)
  }
  l <- stringdist::afind(names, name) # search for similar names
  d <- l$distance |> as.vector()
  ii <- order(d)
  ii <- utils::head(ii, n = n)
  # paste(ii, "=", names[ii])
  names[ii]
}


# name: layout name
layout_exists <- function(x, name, must_exist = FALSE) {
  name <- as.character(name)
  layouts <- get_layout_names(x)
  exists <- name %in% layouts
  if (!must_exist) {
    return(exists)
  }
  if (!exists) {
    similar_msg <- NULL
    similar_layouts <- find_similar_names(name, layouts, n = 2)
    if (!is.null(similar_layouts)) {
      similar_msg <- c("x" = "Did you mean {.or {.val {similar_layouts}}}.")
    }
    cli::cli_abort(c(
      "Layout {.val {name}} does not exists.", similar_msg,
      "i" = cli::col_grey("See {.fn layout_summary} for available layouts")
    ), call = NULL)
  }
  exists
}


layout_is_unique <- function(x, name, require_unique = FALSE, check_exists = FALSE) {
  if (check_exists) {
    layout_exists(x, name, must_exist = TRUE)
  }
  df_all <- layout_summary(x)
  df <- subset(df_all, layout == name)
  name_is_unique <- nrow(df) == 1
  if (!name_is_unique && require_unique) {
    cli::cli_abort(c(
      "Layout {.val {name}} is not unique.",
      "x" = "It exists in {.val {length(df$master)}} masters: {.val {df$master}}"
    ))
  }
  name_is_unique
}
