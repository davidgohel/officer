#' @title PowerPoint table to matrix
#' @description Convert the data in an a 'PowerPoint' table
#' to a matrix or all data to a list of matrices.
#' @param x The rpptx object to convert (as created by [officer::read_pptx()])
#' @param ... Ignored
#' @param slide_id The slide number to load from (NA indicates first slide with
#'   a table, NULL indicates all slides and all tables)
#' @param id The table ID to load from (ignored it `is.null(slide_id)`, NA
#'   indicates to load the first table from the `slide_id`)
#' @param span How should col_span/row_span values be handled? `NA` means
#'   to leave the value as `NA`, and `"fill"` means to fill matrix
#'   cells with the value.
#' @return A matrix with the data, or if `slide_id=NULL`, a list of
#'   matrices
#' @examples
#' library(officer)
#' pptx_file <- system.file(package="officer", "doc_examples", "example.pptx")
#' z <- read_pptx(pptx_file)
#' as.matrix(z, slide_id = NULL)
#' @export
as.matrix.rpptx <- function(x, ..., slide_id=NA_integer_, id=NA_character_, span=c(NA_character_, "fill")) {
  span <- match.arg(span)
  d_summary_raw <- officer::pptx_summary(x)
  # filter to just table data
  d_summary_prep_table <- d_summary_raw[d_summary_raw$content_type %in% "table cell", ]
  stopifnot("No tables to extract"=nrow(d_summary_prep_table) > 0)
  d_summary_prep <-
    d_summary_prep_table[
      d_summary_prep_table$row_span > 0 & d_summary_prep_table$col_span > 0,
    ]
  stopifnot("All cells had col_span=0 or row_span=0"=nrow(d_summary_prep) > 0)
  if (is.null(slide_id)) {
    data_extract <- d_summary_prep
  } else {
    if (is.na(slide_id)) {
      slide_id <- min(d_summary_prep$slide_id)
      id <- min(d_summary_prep$id[d_summary_prep$slide_id == slide_id])
    }
    data_extract <- d_summary_prep[d_summary_prep$slide_id == slide_id & d_summary_prep$id == id, ]
    if (nrow(data_extract) == 0) {
      stop("No data in slide_id=%s, id=%s", as.character(slide_id), as.character(id))
    }
  }
  ret <- list()
  for (current_slide_id in unique(data_extract$slide_id)) {
    ret[[as.character(current_slide_id)]] <-
      as_matrix_rpptx_single_slide(data_extract[data_extract$slide_id == current_slide_id, ], span=span)
  }
  if (length(slide_id) == 1 && length(id) == 1) {
    # When the user is specific, return a matrix (otherwise, return the list of
    # lists of matrices)
    ret <- ret[[1]][[1]]
  }
  ret
}

as_matrix_rpptx_single_slide <- function(x, span=c(NA_character_, "fill")) {
  span <- match.arg(span)
  stopifnot(nrow(x) > 0)
  stopifnot(length(unique(x$slide_id)) == 1)
  ret <- list()
  for (current_id in unique(x$id)) {
    ret[[as.character(current_id)]] <-
      as_matrix_rpptx_single_table(x[x$id == current_id, ], span=span)
  }
  ret
}

as_matrix_rpptx_single_table <- function(x, span=c(NA_character_, "fill")) {
  span <- match.arg(span)
  stopifnot(nrow(x) > 0)
  stopifnot(length(unique(x$slide_id)) == 1)
  stopifnot(length(unique(x$id)) == 1)
  ret <- matrix(NA_character_, nrow=max(x$row_id), ncol=max(x$cell_id))
  for (current_data_idx in seq_len(nrow(x))) {
    if (span %in% "fill") {
      # fill the row/col span
      all_rows <- seq(x$row_id[current_data_idx], x$row_id[current_data_idx] + x$row_span[current_data_idx] - 1)
      all_cols <- seq(x$cell_id[current_data_idx], x$cell_id[current_data_idx] + x$col_span[current_data_idx] - 1)
    } else {
      # just the current cell
      all_rows <- x$row_id[current_data_idx]
      all_cols <- x$cell_id[current_data_idx]
    }
    for (current_col in all_cols) {
      # A loop or some other multiple assignment is needed (multiple indexing
      # does not work on rows and columns simultaneously for matrices)
      ret[all_rows, current_col] <- x$text[current_data_idx]
    }
  }
  ret
}
