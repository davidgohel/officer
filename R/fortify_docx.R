# Replace non breaking hyphen with a text node containing a hyphen-minus = \u002D
replace_no_break_hyphen <- function(x) {
  xml2::xml_replace(
    xml2::xml_find_all(x, ".//w:noBreakHyphen"),
    "w:t",
    "\u002D"
  )
}

unfold_row_wml <- function(node, row_id, preserve = FALSE) {
  is_header_1 <- !inherits(xml_child(node, "w:trPr/w:tblHeader"), "xml_missing")
  is_header_2 <- !inherits(xml_child(node, "w:trPr/w:cnfStyle[@w:firstRow='1']"), "xml_missing")
  is_header <- is_header_1 | is_header_2
  children_ <- xml_children(node)
  cell_nodes <- children_[sapply(children_, function(x) xml_name(x) == "tc")]

  if (preserve) {
    txt <- sapply(cell_nodes, function(x) {
      replace_no_break_hyphen(x)
      paras <- xml2::xml_text(xml2::xml_find_all(x, "./w:p"))
      paste0(paras, collapse = "\n")
    })
  } else {
    txt <- sapply(cell_nodes, function(x) {
      replace_no_break_hyphen(x)
      xml2::xml_text(x)
    })
  }

  col_span <- sapply(cell_nodes, function(x) {
    gs <- xml_child(x, "w:tcPr/w:gridSpan")
    as.integer(xml_attr(gs, "val"))
  })
  col_span[is.na(col_span)] <- 1

  row_span <- lapply(cell_nodes, function(x) {
    node_ <- xml_child(x, "w:tcPr/w:vMerge")
    data.frame(
      row_merge = !inherits(node_, "xml_missing"),
      first = xml_attr(node_, "val") %in% "restart",
      row_span = 1,
      stringsAsFactors = FALSE
    )
  })
  row_span <- rbind_match_columns(row_span)
  txt[row_span$row_merge & !row_span$first] <- NA

  out <- data.frame(
    row_id = row_id, is_header = is_header,
    cell_id = 1 + simple_lag(cumsum(col_span), default = 0),
    text = txt, col_span = col_span, stringsAsFactors = FALSE
  )
  out <- cbind(out, row_span)

  out_add_ <- out[out$col_span > 1, ]

  out_add_ <- mapply(
    function(row_id, is_header, cell_id, text, col_span, row_merge, first, row_span) {
      reps_ <- col_span - 1
      row_num_ <- rep(row_id, reps_)
      is_header_ <- rep(is_header, reps_)
      cell_id_ <- rep(cell_id, reps_)
      text_ <- rep(NA, reps_)
      col_span_ <- rep(0, reps_)
      row_merge_ <- rep(row_merge, reps_)
      first_ <- rep(first, reps_)
      row_span_ <- rep(row_span, reps_)
      out <- data.frame(
        row_id = row_num_, is_header = is_header_, cell_id = cell_id_,
        text = text_, col_span = col_span_, row_merge = row_merge_,
        first = first_, row_span = row_span_,
        stringsAsFactors = FALSE
      )
      out$cell_id <- seq_len(reps_) + cell_id
      out
    }, out_add_$row_id, out_add_$is_header, out_add_$cell_id, out_add_$text,
    out_add_$col_span, out_add_$row_merge, out_add_$first, out_add_$row_span,
    SIMPLIFY = FALSE
  )
  if (length(out_add_) > 0) {
    out_add_ <- rbind_match_columns(out_add_)
    out <- rbind(out, out_add_)
  }
  out[order(out$cell_id), ]
}


docxtable_as_tibble <- function(node, styles, preserve = FALSE) {
  xpath_ <- paste0(xml_path(node), "/w:tr")
  rows <- xml_find_all(node, xpath_)
  if (length(rows) < 1) {
    return(NULL)
  }

  row_details <- mapply(unfold_row_wml, rows, seq_along(rows), preserve = preserve, SIMPLIFY = FALSE)
  row_details <- rbind_match_columns(row_details)
  row_details <- set_row_span(row_details)

  style_node <- xml_child(node, "w:tblPr/w:tblStyle")
  if (inherits(style_node, "xml_missing")) {
    style_name <- NA
  } else {
    style_id <- xml_attr(style_node, "val")
    officer_style_name <- xml_attr(style_node, "tstlname")
    if (!is.na(style_id)) {
      row_details$style_name <- rep(styles$style_name[styles$style_id %in% style_id], nrow(row_details))
    } else if (!is.na(officer_style_name)) {
      row_details$style_name <- rep(officer_style_name, nrow(row_details))
    } else {
      row_details$style_name <- rep(NA_character_, nrow(row_details))
    }
  }
  row_details$content_type <- rep("table cell", nrow(row_details))
  row_details$text[row_details$col_span < 1 | row_details$row_span < 1] <- NA_character_
  row_details
}

#' @importFrom xml2 xml_has_attr
par_as_tibble <- function(node, styles, detailed = FALSE) {
  style_node <- xml_child(node, "w:pPr/w:pStyle")
  if (inherits(style_node, "xml_missing")) {
    style_name <- NA
  } else if (xml_has_attr(style_node, "stlname")) {
    style_name <- xml_attr(style_node, "stlname")
  } else {
    style_id <- xml_attr(style_node, "val")
    style_name <- styles$style_name[styles$style_id %in% style_id]
    if (length(style_name) < 1L) {
      style_name <- NA_character_
    }
  }

  replace_no_break_hyphen(node)

  par_data <- data.frame(
    level = as.integer(xml_attr(xml_child(node, "w:pPr/w:numPr/w:ilvl"), "val")) + 1,
    num_id = as.integer(xml_attr(xml_child(node, "w:pPr/w:numPr/w:numId"), "val")),
    text = xml_text(node),
    style_name = style_name,
    stringsAsFactors = FALSE
  )

  if (detailed) {
    nodes_run <- xml_find_all(node, "w:r")
    run_data <- lapply(nodes_run, run_as_tibble, styles = styles)

    run_data <- mapply(function(x, id) {
      x$id <- id
      x
    }, run_data, seq_along(run_data), SIMPLIFY = FALSE)
    run_data <- rbind_match_columns(run_data)

    par_data$run <- I(list(run_data))
  }

  par_data$content_type <- rep("paragraph", nrow(par_data))
  par_data
}
#' @importFrom xml2 xml_has_attr
val_child <- function(node, child_path, attr = "val", default = NULL) {
  child_node <- xml_child(node, child_path)
  if (inherits(child_node, "xml_missing")) return(NA_character_)
  if (!xml_has_attr(child_node, attr)) default
  else xml_attr(child_node, attr)
}

val_child_lgl <- function(node, child_path, attr = "val", default = NULL) {
  val <- val_child(node = node, child_path = child_path, attr = attr, default = default)
  if (is.na(val)) return(NA)
  else (val %in% c("1", "on", "true"))
}

val_child_int <- function(node, child_path, attr = "val", default = NULL) {
  as.integer(
    val_child(node = node, child_path = child_path, attr = attr, default = default)
  )
}

run_as_tibble <- function(node, styles) {
  style_node <- xml_child(node, "w:rPr/w:rStyle")
  if (inherits(style_node, "xml_missing")) {
    style_name <- NA
  } else {
    style_id <- xml_attr(style_node, "val")
    style_name <- styles$style_name[styles$style_id %in% style_id]
  }
  if (length(style_name) < 1L) {
    style_name <- NA_character_
  }
  run_data <- data.frame(
    text = xml_text(node),
    bold = val_child_lgl(node, "w:rPr/w:b", default = "true"),
    italic = val_child_lgl(node, "w:rPr/w:i", default = "true"),
    underline = val_child(node, "w:rPr/w:u"),
    sz = val_child_int(node, "w:rPr/w:sz"),
    szCs = val_child_int(node, "w:rPr/w:szCs"),
    color = val_child(node, "w:rPr/w:color"),
    shading = val_child(node, "w:rPr/w:shd"),
    shading_color = val_child(node, "w:rPr/w:shd", attr = "color"),
    shading_fill = val_child(node, "w:rPr/w:shd", attr = "fill"),
    style_name = style_name,
    stringsAsFactors = FALSE
  )

  run_data
}

node_content <- function(node, x, preserve = FALSE, detailed = FALSE) {
  node_name <- xml_name(node)
  switch(node_name,
    p = par_as_tibble(node, styles_info(x), detailed = detailed),
    tbl = docxtable_as_tibble(node, styles_info(x), preserve = preserve),
    NULL
  )
}


#' @title Get Word content in a data.frame
#' @description read content of a Word document and
#' return a data.frame representing the document.
#' @note
#' Documents included with [body_add_docx()] will
#' not be accessible in the results.
#' @param x an rdocx object
#' @param preserve If `FALSE` (default), text in table cells is collapsed into a
#'   single line. If `TRUE`, line breaks in table cells are preserved as a "\\n"
#'   character. This feature is adapted from `docxtractr::docx_extract_tbl()`
#'   published under a [MIT
#'   licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE) in
#'   the 'docxtractr' package by Bob Rudis.
#' @param remove_fields if TRUE, prevent field codes from appearing in the
#' returned data.frame.
#' @param detailed Should information on runs be included in summary dataframe?
#'   Defaults to `FALSE`. If `TRUE` a list column `run` is added to the summary
#'   containing a summary of formatting properties of runs as a dataframe with
#'   rows corresponding to a single run and columns containing the information
#'   on formatting properties.
#'
#' @examples
#' example_docx <- system.file(
#'   package = "officer",
#'   "doc_examples/example.docx"
#' )
#' doc <- read_docx(example_docx)
#'
#' docx_summary(doc)
#'
#' docx_summary(doc, preserve = TRUE)[28, ]
#' @export
docx_summary <- function(x, preserve = FALSE, remove_fields = FALSE, detailed = FALSE) {
  if (remove_fields) {
    instrText_nodes <- xml_find_all(x$doc_obj$get(), "//w:instrText")
    xml_remove(instrText_nodes)

    fldData_nodes <- xml_find_all(x$doc_obj$get(), "//w:fldData")
    xml_remove(fldData_nodes)
  }

  all_nodes <- xml_find_all(x$doc_obj$get(), "/w:document/w:body/*[self::w:p or self::w:tbl]")


  data <- lapply(all_nodes, node_content, x = x, preserve = preserve, detailed = detailed)

  data <- mapply(function(x, id) {
    x$doc_index <- id
    x
  }, data, seq_along(data), SIMPLIFY = FALSE)

  data <- rbind_match_columns(data)

  colnames <- c(
    "doc_index", "content_type", "style_name", "text",
    "level", "num_id", "row_id", "is_header", "cell_id",
    "col_span", "row_span", "run"
  )
  colnames <- intersect(colnames, names(data))
  data[, colnames]
}
