#' Change ph labels in a layout
#'
#' There are two versions of the function. The first takes a set of key-value pairs to rename the
#' ph labels. The second uses a right hand side (rhs) assignment to specify the new ph labels.
#' See section *Details*. \cr\cr
#' _NB:_ You can also rename ph labels directly in PowerPoint. Open the master template view
#' (`Alt` + `F10`) and go to `Home` > `Arrange` > `Selection Pane`.
#'
#' @details
#' * Note the difference between the terms `id` and `index`. Both can be found in the output of
#' [layout_properties()]. The unique ph `id` is found in column `id`. The `index` refers to the
#' index of the data frame row.
#' * In a right hand side (rhs) label assignment (`<- new_labels`), there are two ways to
#' optionally specify a subset of phs to rename. In both cases, the length of the rhs vector
#' (the new labels) must match the length of the id or index:
#'   1. use the `id` argument to specify ph ids to rename: `layout_rename_ph_labels(..., id = 2:3) <- new_labels`
#'   2. use an `index` in squared brackets: `layout_rename_ph_labels(...)[1:2] <- new_labels`
#'
#' @export
#' @rdname layout_rename_ph_labels
#' @param x An `rpptx` object.
#' @param layout Layout name or index. Index is the row index of [layout_summary()].
#' @param master Name of master. Only required if the layout name is not unique across masters.
#' @param ... Comma separated list of key-value pairs to rename phs. Either reference a ph via its label
#' (`"old label"` = `"new label"`) or its unique id (`"id"` = `"new label"`).
#' @param .dots Provide a named list or vector of key-value pairs to rename phs
#' (`list("old label"` = `"new label"`).
#' @param id Unique placeholder id (see column `id` in [layout_properties()] or [plot_layout_properties()]).
#' @param value Not relevant for user. A pure technical necessity for rhs assignments.
#' @return Vector of renamed ph labels.
#' @example inst/examples/example_layout_rename_ph_labels.R
#'
layout_rename_ph_labels <- function(x, layout, master = NULL, ..., .dots = NULL) {
  stop_if_not_rpptx(x, "x")
  dots <- list(...)
  dots <- c(dots, .dots)
  if (length(dots) > 0 && !is_named(dots)) {
    cli::cli_abort(
      c("Unnamed arguments are not allowed.",
        "x" = "Arguments {.arg ...} and {.arg .dots} both require key value pairs."
      ),
      call = NULL
    )
  }

  l <- get_layout(x, layout, master)
  lp <- layout_properties(x, l$layout_name, l$master_name)
  if (length(dots) == 0) {
    return(lp$ph_label)
  }
  df_renames <- .rename_df_from_dots(lp, dots)
  .set_ph_labels(l, df_renames)
  reload_slidelayouts(x)

  lp <- layout_properties(x, l$layout_name, l$master_name)
  invisible(lp$ph_label)
}


#' @export
#' @rdname layout_rename_ph_labels
`layout_rename_ph_labels<-` <- function(x, layout, master = NULL, id = NULL, value) {
  l <- get_layout(x, layout, master)
  lp <- layout_properties(x, l$layout_name, l$master_name)

  if (!is.null(id)) {
    if (length(id) != length(value)) {
      cli::cli_abort(
        c("{.arg id} and rhs vector must have the same length",
          "x" = "Number of ids ({.val {length(id)}}) and assigned values ({.val {length(value)}}) differ"
        )
      )
    }
    wrong_ids <- setdiff(id, lp$id)
    n_wrong <- length(wrong_ids)
    if (n_wrong > 0) {
      cli::cli_abort(c(
        "{cli::qty(n_wrong)} {?This/These} id{?s} {?does/do} not exist: {.val {wrong_ids}}",
        "x" = "Choose one of: {.val {lp$id}}",
        "i" = cli::col_grey("Also see {.code plot_layout_properties(..., '{l$layout_name}', '{l$master_nam}')}")
      ))
    }
    .idx <- match(id, lp$id) # user might enter ids in arbitrary order
    lp$ph_label[.idx] <- value
    value <- lp$ph_label
  }
  names(value) <- lp$id
  df_renames <- .rename_df_from_dots(lp, value)
  .set_ph_labels(l, df_renames)
  reload_slidelayouts(x)
}


# heuristic: if a number, then treat as ph_id
.detect_ph_id <- function(x) {
  suppressWarnings({ # avoid character to NA warning
    !is.na(as.numeric(x)) # nchar(x) == 1 &
  })
}


# create data frame with: ph_id, ph_label, ph_label_new as a
# basis for subsequent renaming
#
# CAVEAT: the ph order in layout_properties() (i.e. get_xfrm_data()) is reference for the user.
# Using the 'slide_layout' object's xfrm() method does not yield the same ph order!
# We need to guarantee a proper match here.
#
.rename_df_from_dots <- function(lp, dots) {
  lp <- lp[, c("id", "ph_label")]
  label_old <- names(dots)
  label_new <- as.character(dots)
  is_id <- .detect_ph_id(label_old)
  is_label <- !is_id

  # warn if renaming a duplicate label
  ii <- duplicated(lp$ph_label)
  dupes <- lp$ph_label[ii]
  dupes_used <- intersect(label_old, dupes)
  n_dupes_used <- length(dupes_used)
  if (n_dupes_used > 0) {
    cli::cli_warn(c(
      "When renaming a label with duplicates, only the first occurrence is renamed.",
      "x" = "Renaming {n_dupes_used} ph label{?s} with duplicates: {.val {dupes_used}}"
    ), call = NULL)
  }

  # check for duplicate renames
  is_dupe <- duplicated(label_old)
  if (any(is_dupe)) {
    dupes <- unique(label_old[is_dupe])
    n_dupes <- length(dupes)
    cli::cli_abort(c(
      "Each id or label must only have one rename entry only.",
      "x" = "Found {n_dupes} duplicate id{?s}/label{?s} to rename: {.val {dupes}}"
    ), call = NULL)
  }

  # match by label and check for unknown labels
  label_old_ <- label_old[is_label]
  row_idx_label <- match(label_old_, table = lp$ph_label)
  i_wrong <- is.na(row_idx_label)
  n_wrong <- sum(i_wrong)
  if (n_wrong > 0) {
    cli::cli_abort(c(
      "Can't rename labels that don't exist.",
      "x" = "{cli::qty(n_wrong)}{?This label does/These labels do} not exist: {.val {label_old_[i_wrong]}}"
    ), call = NULL)
  }

  # match by id and check for unknown ids
  id_old_ <- label_old[is_id]
  row_idx_id <- match(id_old_, table = lp$id)
  i_wrong <- is.na(row_idx_id)
  n_wrong <- sum(i_wrong)
  if (n_wrong > 0) {
    cli::cli_abort(c(
      "Can't rename ids that don't exist.",
      "x" = "{cli::qty(n_wrong)}{?This id does/These ids do} not exist: {.val {id_old_[i_wrong]}}"
    ), call = NULL)
  }

  # check for collision between label and id
  idx_collision <- intersect(row_idx_label, row_idx_id)
  n_collision <- length(idx_collision)
  if (n_collision > 0) {
    df <- lp[idx_collision, c("id", "ph_label")]
    pairs <- paste(df$ph_label, "<-->", df$id)
    cli::cli_abort(c(
      "Either specify the label {.emph OR} the id of the ph to rename, not both.",
      "x" = "These labels and ids collide: {.val {pairs}}"
    ), call = NULL)
  }

  lp$ph_label_new <- NA
  lp$ph_label_new[row_idx_label] <- label_new[is_label]
  lp$ph_label_new[row_idx_id] <- label_new[is_id]
  lp[!is.na(lp$ph_label_new), , drop = FALSE]
}


.set_ph_labels <- function(l, df_renames) {
  if (!inherits(l, "layout_info")) {
    cli::cli_abort(
      c("{.arg l} must a a {.cls layout_info} object",
        "x" = "Got {.cls {class(l)[1]}} instead"
      ),
      call = NULL
    )
  }
  layout_xml <- l$slide_layout$get()
  for (i in seq_len(nrow(df_renames))) {
    cnvpr_node <- xml2::xml_find_first(layout_xml, sprintf("p:cSld/p:spTree/*/p:nvSpPr/p:cNvPr[@id='%s']", df_renames$id[i]))
    xml2::xml_set_attr(cnvpr_node, "name", df_renames$ph_label_new[i])
  }
  l$slide_layout$save() # persist changes in slide layout xml file
}
