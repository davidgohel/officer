#' Detect and handle duplicate placeholder labels
#'
#' PowerPoint does not enforce unique placeholder labels in a layout.
#' Selecting a placeholder via its label using [ph_location_label] will throw
#' an error, if the label is not unique. [layout_dedupe_ph_labels] helps to detect,
#' rename, or delete duplicate placholder labels.
#'
#' @param x An `rpptx` object.
#' @param action Action to perform on duplicate placeholder labels. One of:
#' * `detect` (default) = show info on dupes only, make no changes
#' * `rename` = create unique labels. Labels are renamed by appending a sequential number
#'    separated by dot to duplicate labels. For example, `c("title", "title")` becomes `c("title.1", "title.2")`.
#' * `delete` = only keep one of the placeholders with a duplicate label
#' @param print_info Print action information (e.g. renamed placeholders) to console?
#'  Default is `FALSE`. Always `TRUE` for action `detect`.
#' @return A `rpptx` object (with modified placeholder labels).
#' @export
#' @examples
#' x <- read_pptx()
#' layout_dedupe_ph_labels(x)
#'
#' file <- system.file("doc_examples", "ph_dupes.pptx", package = "officer")
#' x <- read_pptx(file)
#' layout_dedupe_ph_labels(x)
#' layout_dedupe_ph_labels(x, "rename", print_info = TRUE)
#'
layout_dedupe_ph_labels <- function(x, action = "detect", print_info = FALSE) {
  if (!inherits(x, "rpptx")) {
    stop("'x' must be an 'rpptx' object", call. = FALSE)
  }
  action <- match.arg(action, c("detect", "rename", "delete"))
  layout_names <- x$slideLayouts$get_metadata()$filename
  xfrm_list <- lapply(layout_names, .dedupe_phs_in_layout, x = x, action = action)
  x <- reload_slidelayouts(x) # reinit slideLayouts to get processed ph labels [e.g. when calling x$slideLayouts$get_xfrm_data()]
  if (print_info | action == "detect") {
    .print_dedupe_info(x = x, xfrm_list = xfrm_list, action = action)
  }
  invisible(x)
}


# handle placeholder labels in a single layout
#
# layout_file: layout filename (e.g. "slideLayout1.xml").
# x: An `rpptx` object
#
# returns: Dataframe with placeholder info. Only needed for .print_dedupe_info()
.dedupe_phs_in_layout <- function(layout_file, x, action = "rename") {
  ph_label <- NULL
  if (!grepl("\\.xml$", layout_file, ignore.case = TRUE)) {
    stop("'layout_file' must be an .xml file", call. = FALSE)
  }
  action <- match.arg(action, c("detect", "rename", "delete"))
  layout <- x$slideLayouts$collection_get(layout_file)
  xfrm <- layout$xfrm()
  xfrm <- subset(xfrm, duplicated(ph_label) | duplicated(ph_label, fromLast = TRUE))
  if (nrow(xfrm) == 0) {
    return()
  }
  xfrm <- transform(xfrm, ph_label_new = make_strings_unique(ph_label), delete_flag = duplicated(ph_label)) # prepare once for all action types
  if (action == "detect") {
    return(xfrm) # no further action required
  } else if (action == "rename") {
    xfrm$delete_flag <- FALSE
  } else if (action == "delete") {
    xfrm$ph_label_new <- xfrm$ph_label
  }

  # rename label or delete ph shape
  layout_xml <- layout$get()
  for (i in 1L:nrow(xfrm)) {
    shape <- xml2::xml_find_first(layout_xml, sprintf("p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr[@id='%s']]", xfrm$id[i]))
    if (xfrm$delete_flag[i]) {
      xml2::xml_remove(shape)
    } else {
      nodes <- xml2::xml_find_first(shape, ".//p:cNvPr")
      xml2::xml_set_attr(nodes, "name", xfrm$ph_label_new[i])
    }
  }
  layout$save() # persist changes in slideout xml file
  xfrm
}


# reload slideLayouts (if layout XML in package_dir has changed)
reload_slidelayouts <- function(x) {
  x$slideLayouts$initialize(x$package_dir,
    master_metadata = x$masterLayouts$get_metadata(),
    master_xfrm = x$masterLayouts$xfrm()
  )
  x
}


# Create unique string by appending a sepatator and a number
# make_strings_unique(c("A", "B", "B", "C", "A"))
make_strings_unique <- function(x, sep = ".") {
  ii <- stats::ave(x, x, FUN = seq_along)
  paste0(x, sep, ii)
}


# helper mostly for testing
has_ph_dupes <- function(x) {
  if (!inherits(x, "rpptx")) {
    stop("'x' must be an 'rpptx' object", call. = FALSE)
  }
  xfrm <- x$slideLayouts$get_xfrm_data()
  dupes <- stats::aggregate(ph_label ~ master_name + name, data = xfrm, FUN = function(x) sum(duplicated(x)) > 0)
  any(dupes$ph_label)
}


# print info on what was done (if print_info = TRUE)
.print_dedupe_info <- function(x, xfrm_list, action) {
  .df_1 <- do.call(rbind, xfrm_list)
  if (is.null(.df_1)) {
    cat("No duplicate placeholder labels detected.")
    return(invisible(NULL))
  }
  .df_2 <- x$slideLayouts$get_xfrm_data()
  .df_2 <- unique(.df_2[, c("master_file", "master_name"), drop = FALSE])
  df <- merge(.df_1, .df_2, sort = FALSE)
  rownames(df) <- NULL
  df <- df[, c("master_name", "name", "ph_label", "ph_label_new", "delete_flag"), drop = FALSE]
  colnames(df)[2] <- "layout_name"
  if (action == "detect") {
    cat("Placeholders with duplicate labels:\n")
    cat(cli::col_grey("* 'ph_label_new' = new placeholder label for action = 'rename'\n"))
    cat(cli::col_grey("* 'delete_flag' = deleted placeholders for action = 'delete'\n"))
  } else if (action == "rename") {
    df$delete_flag <- NULL
    cat("Renamed duplicate placeholder labels:\n")
    cat(cli::col_grey("* 'ph_label_new' = new placeholder label\n"))
  } else if (action == "delete") {
    df <- df[df$delete_flag, , drop = FALSE]
    df$ph_label_new <- NULL
    cat("Removed placeholders with duplicate labels:\n")
  }
  print(df)
}
