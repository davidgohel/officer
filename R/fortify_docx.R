#' @importFrom cli
#'  cli_abort
#' @importFrom dplyr
#'  mutate case_when left_join arrange across all_of lag consecutive_id syms n
#'  select inner_join if_else filter relocate summarise group_by as_tibble first
#'  bind_rows
#' @importFrom tidyr
#'  complete
#' @importFrom xml2
#'  as_xml_document xml_has_attr xml_child xml_name xml_children xml_replace

# utils ----------
extract_runs_attr_str <- function(
    run_nodes,
    node_xpath,
    attr_name,
    default = NULL
) {
  nodes <- suppressWarnings({
    xml_child(run_nodes, node_xpath)
  })
  nodes_exist <- !is.na(xml_name(nodes))

  str_values <- rep("", length(run_nodes))
  tmp_results <- xml_attr(nodes, attr_name)
  str_values[nodes_exist] <- tmp_results[nodes_exist]
  if (!is.null(default)) {
    str_values[is.na(str_values)] <- default
  }

  str_values[str_values %in% ""] <- NA_character_
  str_values
}

extract_runs_attr_lgl <- function(
    run_nodes,
    node_xpath,
    attr_name,
    default = NULL,
    true_values = c("1", "on", "true", NA_character_)
) {
  nodes <- xml_child(run_nodes, node_xpath)
  nodes_exist <- !is.na(xml_name(nodes))

  lgl_values <- rep(default, length(run_nodes))
  lgl_values[nodes_exist] <- xml_attr(nodes[nodes_exist], attr_name) %in%
    true_values
  lgl_values
}

add_text_column <- function(x) {
  mutate(
    .data = x,
    text = case_when(
      !is.na(.data$run_content_text) ~ .data$run_content_text,
      .default = ""
    )
  )
}

#' @title Compute tables information
#' @description Joins and processes table-related information data frames
#'   to compute cell spans and merges.
#' @param tmp_infotbl_tablecells_paragraphs data.frame with paragraph
#'   information in table cells
#' @param tmp_infotbl_tablecells data.frame with table cells information
#' @param tmp_infotbl_tablerows data.frame with table rows information
#' @param tmp_infotbl_tables data.frame with tables information
#' @return A data.frame with complete table information including computed
#'   cell_id, row_span, and col_span values.
#' @noRd
infotbl_tables_compute <- function(
  tmp_infotbl_tablecells_paragraphs,
  tmp_infotbl_tablecells,
  tmp_infotbl_tablerows,
  tmp_infotbl_tables
) {
  infotbl_tables <- left_join(
    tmp_infotbl_tablecells_paragraphs,
    tmp_infotbl_tablecells,
    by = "cell_id"
  )
  infotbl_tables <- left_join(
    infotbl_tables,
    tmp_infotbl_tablerows,
    by = c("table_index", "row_id")
  )
  infotbl_tables <- left_join(
    infotbl_tables,
    tmp_infotbl_tables,
    by = c("table_index")
  )
  infotbl_tables <- arrange(
    .data = infotbl_tables,
    across(all_of(
      c("table_index", "row_id", "cell_id", "doc_index")
    ))
  )
  infotbl_tables <- mutate(
    .data = infotbl_tables,
    .by = all_of(c("table_index", "row_id")),
    add_span = as.integer(
      lag(.data$col_span, default = "1")
    ) - 1L,
    cell_id = consecutive_id(.data$cell_id) + cumsum(.data$add_span),
    add_span = NULL
  )
  infotbl_tables <- tidyr::complete(
    infotbl_tables,
    !!!syms(c("table_index", "row_id", "cell_id")),
    fill = list(
      row_merge = FALSE,
      first = FALSE,
      col_span = "0",
      is_header = FALSE
    )
  )

  infotbl_tables <- mutate(
      .data = infotbl_tables,
      .by = all_of(c("table_index", "cell_id")),
      merge_group = cumsum(.data$first) + cumsum(!.data$row_merge)
    )

  infotbl_tables <- mutate(
    .data = infotbl_tables,
    .by = all_of(c("table_index", "cell_id", "merge_group")),
    row_span = n(),
    .after = all_of("col_span")
  )
  infotbl_tables <- mutate(
    .data = infotbl_tables,
    merge_group = NULL,
    row_span = case_when(
      .data$row_merge & !.data$first ~ 0L,
      .default = .data$row_span
    )
  )

  infotbl_tables
}

#' @title Augment data with non-text content
#' @description Adds information about images, footnotes, and hyperlinks
#'   to the data by joining with document relationships and processing
#'   footnote XML content.
#' @param data data.frame with document content including IDs for
#'   images (blip_id), footnotes (footnote_id), and hyperlinks (hyperlink_id)
#' @param x an rdocx object containing document relationships and footnotes
#' @return A data.frame with image paths, footnote text, and hyperlink URLs
#'   added, and temporary ID columns removed.
#' @noRd
augment_with_non_text_content <- function(data, x) {
  doc_rel <- x$doc_obj$rel_df()

  # add images
  doc_rel_images <- doc_rel[basename(doc_rel$type) %in% "image", ]
  doc_rel_images <- select(
    doc_rel_images,
    all_of(c("id", "target"))
  )
  names(doc_rel_images) <- c("blip_id", "image_path")
  doc_rel_images$image_path <- file.path(x$package_dir, "word", doc_rel_images$image_path)

  data <- left_join(
    data,
    doc_rel_images,
    by = "blip_id"
  )
  data$blip_id <- NULL

  # add footnotes
  ref_wml_str <- x$footnotes$encode_wml_str(additional_ns = list())
  ref_str <- rep(NA_character_, length(ref_wml_str))
  suppressWarnings({
    ref_str <- vapply(
      ref_wml_str,
      function(x) {
        str <- xml_text(xml2::as_xml_document(x))
        paste0(str, collapse = "\n")
      },
      FUN.VALUE = NA_character_
    )
    ref_str
  })
  ref_data <- data.frame(
    footnote_id = names(ref_str),
    footnote_text = unname(ref_str)
  )
  data <- left_join(
    data,
    ref_data,
    by = "footnote_id"
  )
  data$footnote_id <- NULL

  # add hyperlinks
  run_within_hl_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:hyperlink/w:r"
  )

  hl_nodes <- xml_child(run_within_hl_nodes, "parent::w:hyperlink")
  table_hyperlink <- data.frame(
    hyperlink_id = xml_attr(hl_nodes, "id"),
    link_to_bookmark = xml_attr(hl_nodes, "anchor"),
    run_index = as.integer(xml_attr(run_within_hl_nodes, "run_index"))
  )

  doc_rel_external_hyperlink <- doc_rel[
    basename(doc_rel$type) %in% "hyperlink",
  ]
  doc_rel_external_hyperlink <- select(
    doc_rel_external_hyperlink,
    all_of(c("id", "target"))
  )
  names(doc_rel_external_hyperlink) <- c("hyperlink_id", "link")

  table_to_anchor <- dplyr::anti_join(
    table_hyperlink,
    doc_rel_external_hyperlink,
    by = "hyperlink_id"
  )
  table_to_anchor$hyperlink_id <- NULL
  data <- left_join(
    data,
    table_to_anchor,
    by = c("run_index")
  )

  table_hyperlink <- inner_join(
    table_hyperlink,
    doc_rel_external_hyperlink,
    by = "hyperlink_id"
  )
  data <- left_join(
    data,
    select(
      table_hyperlink,
      all_of(c("hyperlink_id", "run_index", "link"))
    ),
    by = c("hyperlink_id", "run_index")
  )
  data$hyperlink_id <- NULL

  data
}

#' @title Augment data with style names
#' @description Adds style names (character, table, paragraph) to the data
#'   by joining with styles information and removes temporary style ID columns.
#' @param data data.frame with document content including style IDs
#'   (run_stylename, tab_style_name, par_style_name)
#' @param styles_data data.frame with style information containing
#'   style_type, style_id, and style_name columns
#' @return A data.frame with human-readable style names added
#'   (character_stylename, table_stylename, paragraph_stylename) and
#'   temporary style ID columns removed.
#' @noRd
augment_with_styles <- function(data, styles_data) {
  styles_character <- filter(
    styles_data,
    .data$style_type %in% "character"
  )
  styles_character <- select(
    styles_character,
    all_of(c("style_id", "style_name"))
  )
  names(styles_character) <- c("run_stylename", "character_stylename")
  data <- left_join(
    data,
    styles_character,
    by = "run_stylename"
  )

  styles_table <- filter(styles_data, .data$style_type %in% "table")
  styles_table <- select(
    styles_table,
    all_of(c("style_id", "style_name"))
  )
  names(styles_table) <- c("tab_style_name", "table_stylename")
  data <- left_join(
    data,
    styles_table,
    by = "tab_style_name"
  )

  styles_paragraph <- filter(
    styles_data,
    .data$style_type %in% "paragraph"
  )
  styles_paragraph <- select(
    styles_paragraph,
    all_of(c("style_id", "style_name"))
  )
  names(styles_paragraph) <- c("par_style_name", "paragraph_stylename")
  data <- left_join(
    data,
    styles_paragraph,
    by = "par_style_name"
  )
  data <- select(
    data,
    -all_of(c("par_style_name", "tab_style_name", "run_stylename"))
  )

  data
}

#' @title Join information tables for document content
#' @description Joins runs, paragraphs, and tables information data frames
#'   and processes the final document content structure.
#' @param infotbl_runs_contents data.frame with run content information
#'   (text, images, footnotes, etc.)
#' @param infotbl_runs data.frame with run formatting information
#'   (bold, italic, font, etc.)
#' @param infotbl_paragraphs data.frame with paragraph information
#'   (style, alignment, numbering, etc.)
#' @param infotbl_tables data.frame with table information
#'   (cell positions, spans, etc.)
#' @return A data.frame with all information joined and properly ordered,
#'   with consecutive run and content indices.
#' @noRd
infotbl_join <- function(
  infotbl_runs_contents,
  infotbl_runs,
  infotbl_paragraphs,
  infotbl_tables
) {
  data <- left_join(
    infotbl_runs_contents,
    infotbl_runs,
    by = "run_index"
  )
  data <- left_join(
    data,
    infotbl_paragraphs,
    by = "doc_index"
  )

  data <- left_join(
    data,
    infotbl_tables,
    by = "doc_index"
  )

  data <- relocate(
    .data = data,
    all_of(
      c("doc_index", "run_index", "run_content_index",
        "table_index", "row_id", "cell_id")
    )
  )
  data <- mutate(
    .data = data,
    doc_index = as.integer(.data$doc_index),
    run_index = as.integer(.data$run_index)
  )
  data <- arrange(
    .data = data,
    across(all_of(
      c("doc_index", "run_index", "run_content_index")
    ))
  )
  data <- mutate(
    .data = data,
    .by = all_of(c("doc_index", "run_index")),
    .keep = "all",
    run_content_index = consecutive_id(.data$run_content_index)
  )

  data
}


# extract tools ----
docx_tablecells_information <- function(tc_nodes) {

  col_span_str <- rep("1", length(tc_nodes))
  col_span_nodes <- xml_child(tc_nodes, "w:tcPr/w:gridSpan")
  has_col_span_nodes <- xml_name(col_span_nodes) %in% "gridSpan"
  col_span_str[has_col_span_nodes] <- xml_attr(
    col_span_nodes[has_col_span_nodes],
    "val"
  )

  vmerge_nodes <- xml_child(tc_nodes, "w:tcPr/w:vMerge")
  row_merge <- !is.na(xml_name(vmerge_nodes))
  first <- xml_attr(vmerge_nodes, "val") %in% "restart"

  tc_index <- xml_attr(tc_nodes, "tc_index")
  table_index_sel <- xml_child(tc_nodes, "parent::w:tr/parent::w:tbl")
  table_index <- xml_attr(table_index_sel, "table_index")
  tr_index_sel <- xml_child(tc_nodes, "parent::w:tr")
  tr_index <- xml_attr(tr_index_sel, "tr_index")

  data.frame(
    table_index = as.integer(table_index),
    row_id = as.integer(tr_index),
    cell_id = as.integer(tc_index),
    row_merge = row_merge,
    first = first,
    col_span = col_span_str
  )
}

docx_p_cells_information <- function(p_in_cell_nodes) {
  tc_index <- xml_child(p_in_cell_nodes, "parent::w:tc")
  tc_index <- xml_attr(tc_index, "tc_index")
  doc_index <- xml_attr(p_in_cell_nodes, "doc_index")

  data.frame(
    doc_index = as.integer(doc_index),
    cell_id = as.integer(tc_index)
  )
}
docx_p_information <- function(p_nodes) {
  ilvl_nodes <- xml_child(p_nodes, "w:pPr/w:numPr/w:ilvl")
  numId_nodes <- xml_child(p_nodes, "w:pPr/w:numPr/w:numId")
  pstyle_nodes <- xml_child(p_nodes, "w:pPr/w:pStyle")
  jc_nodes <- xml_child(p_nodes, "w:pPr/w:jc")
  kn_nodes <- xml_child(p_nodes, "w:pPr/w:keepNext")

  doc_index <- xml_attr(p_nodes, "doc_index")

  has_num_pr <- xml_name(xml_child(p_nodes, "w:pPr/w:numPr")) %in% "numPr"

  level_str <- rep(NA_character_, length(p_nodes))
  level_str[has_num_pr] <- xml_attr(ilvl_nodes[has_num_pr], "val")

  num_id_str <- rep(NA_character_, length(p_nodes))
  num_id_str[has_num_pr] <- xml_attr(numId_nodes[has_num_pr], "val")

  data.frame(
    doc_index = as.integer(doc_index),
    par_style_name = xml_attr(pstyle_nodes, "val"),
    keep_with_next = xml_name(kn_nodes) %in% "keepNext",
    align = xml_attr(jc_nodes, "val"),
    level = as.integer(level_str) + 1L,
    num_id = as.integer(num_id_str),
    stringsAsFactors = FALSE
  )
}

docx_tables_information <- function(tbl_nodes) {
  tbl_stylename <- xml_attr(tbl_nodes, "val")
  tblstyle_nodes <- xml_child(
    tbl_nodes,
    "w:tblPr/w:tblStyle"
  )
  data.frame(
    tab_style_name = xml_attr(tblstyle_nodes, "val"),
    table_index = as.integer(xml_attr(tbl_nodes, "table_index"))
  )
}

docx_runs_information <- function(run_nodes) {

  bold <- extract_runs_attr_lgl(
    run_nodes = run_nodes,
    node_xpath = "w:rPr/w:b",
    attr_name = "val",
    default = FALSE
  )
  italic <- extract_runs_attr_lgl(
    run_nodes = run_nodes,
    node_xpath = "w:rPr/w:i",
    attr_name = "val",
    default = FALSE
  )
  underline <- extract_runs_attr_lgl(
    run_nodes = run_nodes,
    node_xpath = "w:rPr/w:u",
    attr_name = "val",
    true_values = c(
      "single",
      "words",
      "double",
      "thick",
      "dash",
      NA_character_
    ),
    default = FALSE
  )

  shd_nodes <- xml_child(run_nodes, "w:rPr/w:shd")
  shading <- xml_attr(shd_nodes, "val")
  shading[!is.na(shading)] <- sprintf("#%s", shading[!is.na(shading)])
  shading_color <- xml_attr(shd_nodes, "color")
  shading_color[!is.na(shading_color)] <- sprintf("#%s", shading_color[!is.na(shading_color)])
  shading_fill <- xml_attr(shd_nodes, "fill")
  shading_fill[!is.na(shading_fill)] <- sprintf("#%s", shading_fill[!is.na(shading_fill)])

  run_stylename <- extract_runs_attr_str(
    run_nodes = run_nodes,
    node_xpath = "w:rPr/w:rStyle",
    attr_name = "val",
    default = NA_character_
  )
  hyperlink_id <- extract_runs_attr_str(
    run_nodes = run_nodes,
    node_xpath = "parent::w:hyperlink",
    attr_name = "id",
    default = NA_character_
  )

  run_index <- xml_attr(run_nodes, "run_index")

  rfonts_nodes <- xml_child(run_nodes, "w:rPr/w:rFonts")

  doc_index <- extract_runs_attr_str(
    run_nodes = run_nodes,
    node_xpath = "parent::w:p|parent::*/parent::w:p",
    attr_name = "doc_index",
    default = NA_character_
  )

  sz <- xml_attr(xml_child(run_nodes, "w:rPr/w:sz"), "val")
  sz <- as.integer(sz)

  sz_cs = xml_attr(xml_child(run_nodes, "w:rPr/w:szCs"), "val")
  sz_cs <- as.integer(sz_cs)

  color <- xml_attr(xml_child(run_nodes, "w:rPr/w:color"), "val")
  color[!is.na(color)] <- sprintf("#%s", color[!is.na(color)])

  data.frame(
    doc_index = as.integer(doc_index),
    run_index = run_index,
    run_stylename = run_stylename,
    hyperlink_id = hyperlink_id,
    sz = sz,
    sz_cs = sz_cs,
    font_family_ascii = xml_attr(rfonts_nodes, "ascii"),
    font_family_eastasia = xml_attr(rfonts_nodes, "eastAsia"),
    font_family_hansi = xml_attr(rfonts_nodes, "hAnsi"),
    font_family_cs = xml_attr(rfonts_nodes, "cs"),
    bold = bold,
    italic = italic,
    underline = underline,
    color = color,
    shading = shading,
    shading_color = shading_color,
    shading_fill = shading_fill
  )
}


docx_runs_content_information <- function(run_nodes) {
  run_content_nodes_names <- xml_name(run_nodes)

  suppressWarnings({
    blip_id_ <- xml_attr(xml_child(run_nodes, "wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip"), "embed")
    footnote_id_ <- xml_attr(run_nodes, "id")
    text_values <- xml_text(run_nodes)
  })
  field_code <- case_when(
    run_content_nodes_names %in% "instrText" ~ text_values,
    .default = NA_character_
  )
  data.frame(
    blip_id = case_when(
      run_content_nodes_names %in% "drawing" ~ blip_id_,
      .default = NA_character_
    ),
    footnote_id = case_when(
      run_content_nodes_names %in% "footnoteReference" ~ footnote_id_,
      .default = NA_character_
    ),
    field_code = field_code,
    run_content_text = case_when(
      run_content_nodes_names %in% "t" ~ text_values,
      run_content_nodes_names %in% "noBreakHyphen" ~ "\u002D",
      run_content_nodes_names %in% c("br", "cr") ~ "\n",
      run_content_nodes_names %in% "tab" ~ "\t",
      run_content_nodes_names %in% "instrText" ~ field_code,
      run_content_nodes_names %in% "softHyphen" ~ "",
      run_content_nodes_names %in% "sym" ~ xml_attr(run_nodes, "char"),
      .default = NA_character_
    ),
    run_content_index = xml_attr(run_nodes, "run_content_index"),
    run_index = xml_attr(xml_child(run_nodes, "parent::w:r"), "run_index")
  )
}

docx_tablerows_information <- function(tr_nodes) {
  tbl_header_nodes <- xml_child(tr_nodes, "w:trPr/w:tblHeader")
  first_row_nodes <- xml_child(tr_nodes, "w:trPr/w:cnfStyle[@w:firstRow='1']")
  is_header <- xml_name(first_row_nodes) %in%
    "cnfStyle" |
    xml_name(tbl_header_nodes) %in% "tblHeader"
  tr_index <- xml_attr(tr_nodes, "tr_index")

  table_index <- xml_child(tr_nodes, "parent::w:tbl")
  table_index <- xml_attr(table_index, "table_index")

  data.frame(
    table_index = as.integer(table_index),
    row_id = as.integer(tr_index),
    is_header = is_header
  )
}


# main ----------

#' @title Summarise data as paragraphs
#' @description Summarises detailed run-level data into paragraph-level
#'   summaries by collapsing text within paragraphs. Handles both regular
#'   paragraphs and paragraphs within table cells differently.
#' @param data data.frame with detailed run-level information including
#'   doc_index, table_index, paragraph_stylename, and other formatting details
#' @return A data.frame summarised at the paragraph level with collapsed text
#'   in the column mamed `text`. For tables, includes cell and row
#'   information; for regular paragraphs, includes only doc_index and style.
#' @noRd
summarise_as_paragraph <- function(data, preserve = FALSE) {
  finest_grp_by <- c(
    "doc_index",
    "content_type",
    "table_index",
    "row_id",
    "cell_id",
    "is_header",
    "row_span",
    "col_span",
    "table_stylename",
    "paragraph_stylename"
  )

  table_data <- filter(.data = data, !is.na(.data$table_index))
  table_data <- add_text_column(table_data)
  table_data <- summarise(
      .data = table_data,
      .by = all_of(finest_grp_by),
      text = paste0(.data$text, collapse = "")
    )

  if (preserve) {
    finest_grp_by <- c(
      "content_type",
      "table_index",
      "row_id",
      "cell_id",
      "is_header",
      "row_span",
      "col_span",
      "table_stylename",
      "paragraph_stylename"
    )
    table_data <- summarise(
      .data = table_data,
      doc_index = first(.data$doc_index),
      text = paste0(.data$text, collapse = "\n"),
      .by = all_of(finest_grp_by)
    )
  }

  par_data <- filter(.data = data, is.na(.data$table_index))
  par_data <- add_text_column(par_data)
  par_data <- summarise(
      .data = par_data,
      .by = all_of(
        c("doc_index", "content_type", "paragraph_stylename")
      ),
      text = paste0(.data$text, collapse = "")
    )

  dataset <- bind_rows(par_data, table_data)

  # patch for WordR
  names(dataset)[names(dataset) %in% "paragraph_stylename"] <- "style_name"

  dataset
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
#' @param detailed Should run-level information be included in the dataframe?
#'   Defaults to `FALSE`. If `TRUE`, the dataframe contains detailed information
#'   about each run (text formatting, images, hyperlinks, etc.) instead of
#'   collapsing content at the paragraph level. When `FALSE`, run-level
#'   information such as images, hyperlinks, and text formatting is not available
#'   since data is aggregated at the paragraph level.
#' @return A data.frame with the following columns depending on the value of `detailed`:
#'
#' When `detailed = FALSE` (default), the data.frame contains:
#'
#' - `doc_index`: Document element index (integer).
#' - `content_type`: Type of content: "paragraph" or "table cell" (character).
#' - `style_name`: Name of the paragraph style (character).
#' - `text`: Collapsed text content of the paragraph or cell (character).
#' - `table_index`: Index of the table (integer). `NA` for non-table content.
#' - `row_id`: Row position in table (integer). `NA` for non-table content.
#' - `cell_id`: Cell position in table row (integer). `NA` for non-table content.
#' - `is_header`: Whether the row is a table header (logical). `NA` for non-table content.
#' - `row_span`: Number of rows spanned by the cell (integer). `0` for merged cells. `NA` for non-table content.
#' - `col_span`: Number of columns spanned by the cell (character). `NA` for non-table content.
#' - `table_stylename`: Name of the table style (character). `NA` for non-table content.
#'
#' When `detailed = TRUE`, the data.frame contains additional run-level information:
#'
#' - `run_index`: Index of the run within the paragraph (integer).
#' - `run_content_index`: Index of content element within the run (integer).
#' - `run_content_text`: Text content of the run element (character).
#' - `image_path`: Path to embedded image stored in the temporary directory
#'   associated with the rdocx object (character).
#'   Images should be copied to a permanent location before closing the R
#'   session if needed.
#' - `field_code`: Field code content (character).
#' - `footnote_text`: Footnote text content (character).
#' - `link`: Hyperlink URL (character).
#' - `link_to_bookmark`: Internal bookmark anchor name for hyperlinks (character).
#' - `bookmark_start`: Name of the bookmark starting at this run (character).
#' - `character_stylename`: Name of the character/run style (character).
#' - `sz`: Font size in half-points (integer).
#' - `sz_cs`: Complex script font size in half-points (integer).
#' - `font_family_ascii`: Font family for ASCII characters (character).
#' - `font_family_eastasia`: Font family for East Asian characters (character).
#' - `font_family_hansi`: Font family for high ANSI characters (character).
#' - `font_family_cs`: Font family for complex script characters (character).
#' - `bold`: Whether the run is bold (logical).
#' - `italic`: Whether the run is italic (logical).
#' - `underline`: Whether the run is underlined (logical).
#' - `color`: Text color in hexadecimal format (character).
#' - `shading`: Shading pattern (character).
#' - `shading_color`: Shading foreground color (character).
#' - `shading_fill`: Shading background fill color (character).
#' - `keep_with_next`: Whether paragraph should stay with next (logical).
#' - `align`: Paragraph alignment (character).
#' - `level`: Numbering level (integer). `NA` if not a numbered list.
#' - `num_id`: Numbering definition ID (integer). `NA` if not a numbered list.
#' @examples
#' example_docx <- system.file(
#'   package = "officer",
#'   "doc_examples/example.docx"
#' )
#' doc <- read_docx(example_docx)
#'
#' docx_summary(doc)
#'
#' docx_summary(doc, detailed = TRUE)
#' @export
docx_summary <- function(x, preserve = FALSE, remove_fields = FALSE, detailed = FALSE) {
  if (!requireNamespace("dplyr")) {
    cli::cli_abort(
      "package {.pkg dplyr} is required for running {.fn docx_summary}."
    )
  }
  if (!requireNamespace("tidyr")) {
    cli::cli_abort(
      "package {.pkg tidyr} is required for running {.fn docx_summary}."
    )
  }

  if (remove_fields) {
    instrText_nodes <- xml_find_all(x$doc_obj$get(), "//w:instrText")
    xml_remove(instrText_nodes)

    fldData_nodes <- xml_find_all(x$doc_obj$get(), "//w:fldData")
    xml_remove(fldData_nodes)
  }

  # add info to xml -----
  ## p_nodes
  p_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:p"
  )
  xml_attr(p_nodes, "w:doc_index") <- seq_along(p_nodes)

  ## tbl_nodes
  tbl_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:tbl"
  )
  xml_attr(tbl_nodes, "w:table_index") <- seq_along(tbl_nodes)

  ## tr_nodes
  tr_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:tr"
  )
  xml_attr(tr_nodes, "w:tr_index") <- seq_along(tr_nodes)

  ## tc_nodes
  tc_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:tc"
  )
  xml_attr(tc_nodes, "w:tc_index") <- seq_along(tc_nodes)

  ## run_nodes
  run_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:r"
  )
  xml_attr(run_nodes, "w:run_index") <- seq_along(run_nodes)

  ## run_content_nodes
  run_content_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:r/*[not(self::w:rPr)]"
  )
  xml_attr(run_content_nodes, "w:run_content_index") <- seq_along(run_content_nodes)

  ## bookmark_nodes
  bookmark_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:bookmarkStart[following-sibling::w:r]"
  )
  bookmark_nodes_siblings <- xml_find_all(
    x$doc_obj$get(),
    "//w:bookmarkStart/following-sibling::w:r"
  )
  data_bookmark <- data.frame(
    bookmark_start = xml_attr(bookmark_nodes, "name"),
    run_index = xml_attr(bookmark_nodes_siblings, "run_index")
  )

  ## p_in_cell_nodes
  p_in_cell_nodes <- xml_find_all(
    x$doc_obj$get(),
    "//w:tbl/w:tr/w:tc/w:p"
  )

  # info for runs: infotbl_runs -----
  infotbl_runs_contents <- docx_runs_content_information(run_content_nodes)
  infotbl_runs <- docx_runs_information(run_nodes)
  infotbl_runs <- left_join(infotbl_runs, data_bookmark, by = "run_index")

  # info for tables: infotbl_tables -----
  tmp_infotbl_tables <- docx_tables_information(tbl_nodes)
  tmp_infotbl_tablecells <- docx_tablecells_information(tc_nodes)
  tmp_infotbl_tablecells_paragraphs <- docx_p_cells_information(p_in_cell_nodes)
  tmp_infotbl_tablerows <- docx_tablerows_information(tr_nodes)
  ## final table infotbl_tables
  infotbl_tables <- infotbl_tables_compute(
    tmp_infotbl_tablecells_paragraphs,
    tmp_infotbl_tablecells,
    tmp_infotbl_tablerows,
    tmp_infotbl_tables
  )

  # info for paragraphs: infotbl_paragraphs -----
  infotbl_paragraphs <- docx_p_information(p_nodes)

  # final joins -------
  data <- infotbl_join(infotbl_runs_contents, infotbl_runs, infotbl_paragraphs, infotbl_tables)

  # add stylenames -----
  styles_data <- select(
    styles_info(x),
    all_of(c("style_type", "style_id", "style_name"))
  )
  data <- augment_with_styles(data, styles_data)

  # add non textual content -----
  data <- augment_with_non_text_content(data, x)

  data <- mutate(
    .data = data,
    .by = all_of(c("doc_index")),
    .keep = "all",
    run_index = consecutive_id(.data$run_index)
  )
  data <- mutate(
    .data = data,
    content_type = if_else(
      is.na(.data$table_index),
      "paragraph",
      "table cell"
    )
  )

  column_names <- c("doc_index", "content_type", "run_index", "run_content_index",

    "run_content_text", "image_path", "field_code", "footnote_text",
    "link", "link_to_bookmark", "bookmark_start",
    "character_stylename", "sz", "sz_cs",
    "font_family_ascii", "font_family_eastasia", "font_family_hansi", "font_family_cs",
    "bold", "italic", "underline", "color",
    "shading", "shading_color", "shading_fill",

    "paragraph_stylename", "keep_with_next", "align", "level", "num_id",

    "table_index", "row_id", "cell_id", "col_span", "row_span", "is_header",
    "table_stylename")

  data <- select(.data = data, all_of(column_names))

  if (!detailed) {
    data <- summarise_as_paragraph(data, preserve = preserve)
  }

  data
}
