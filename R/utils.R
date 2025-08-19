#' @importFrom xml2 xml_attr<- xml_name<- xml_text<- as_list as_xml_document read_xml
#' write_xml xml_add_child xml_add_parent xml_add_sibling xml_attr xml_attrs xml_child
#' xml_children xml_find_all xml_find_first xml_length xml_missing xml_name xml_ns
#' xml_path xml_remove xml_replace xml_set_attr xml_set_attrs xml_text
#' @importFrom stats setNames


read_xfrm <- function(nodeset, file, name){

  if( length(nodeset) < 1 ){
    return(data.frame(stringsAsFactors = FALSE, type = character(0),
                   id = character(0),
                   ph_label = character(0),
                   ph = character(0),
                   file = character(0),
                   offx = integer(0),
                   offy = integer(0),
                   cx = integer(0),
                   cy = integer(0),
                   rotation = integer(0),
                   name = character(0),
                   fld_id = character(0),
                   fld_type = character(0)
                   ))
  }

  ph <- xml_child(nodeset, "p:nvSpPr/p:nvPr/p:ph")
  type <- xml_attr(ph, "type")
  type[is.na(type)] <- "body"
  id <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "id")
  label <- xml_attr(xml_child(nodeset, "/p:cNvPr"), "name")

  off <- xml_child(nodeset, "p:spPr/a:xfrm/a:off")
  ext <- xml_child(nodeset, "p:spPr/a:xfrm/a:ext")
  rot <- xml_child(nodeset, "p:spPr/a:xfrm")

  fld_id <- xml_attr(xml_child(nodeset, "/p:txBody/a:p/a:fld"), "id")
  fld_type <- xml_attr(xml_child(nodeset, "/p:txBody/a:p/a:fld"), "type")

  data.frame(stringsAsFactors = FALSE, type = type, id = id,
          ph_label = label,
          ph = as.character(ph),
          file = basename(file),
          offx = as.integer(xml_attr(off, "x")),
          offy = as.integer(xml_attr(off, "y")),
          cx = as.integer(xml_attr(ext, "cx")),
          cy = as.integer(xml_attr(ext, "cy")),
          rotation = as.integer(xml_attr(rot, "rot")),
          fld_id,
          fld_type,
          name = name )
}


fortify_master_xfrm <- function(master_xfrm) {
  master_xfrm <- as.data.frame(master_xfrm)
  has_type <- grepl("type=", master_xfrm$ph)
  master_xfrm <- master_xfrm[has_type, ]
  if (nrow(master_xfrm) > 0) {    # see #597
    list_xfrm <- split(master_xfrm, master_xfrm$file)
    list_xfrm <- lapply(list_xfrm, function(x) {
      x[!duplicated(x$type), , drop = FALSE]
    })
    master_xfrm <- do.call("rbind", list_xfrm)
  }

  tmp_names <- names(master_xfrm)

  old_ <- c("offx", "offy", "cx", "cy", "fld_id", "fld_type", "name")
  new_ <- c("offx_ref", "offy_ref", "cx_ref", "cy_ref", "fld_id_ref", "fld_type_ref", "master_name")
  tmp_names[match(old_, tmp_names)] <- new_
  names(master_xfrm) <- tmp_names
  master_xfrm$id <- NULL
  master_xfrm$ph <- NULL
  master_xfrm$ph_label <- NULL
  master_xfrm$rotation <- NULL

  master_xfrm
}


xfrmize <- function(slide_xfrm, master_xfrm) {
  x <- as.data.frame(slide_xfrm)

  master_ref <- unique(data.frame(
    file = master_xfrm$file,
    master_name = master_xfrm$name,
    stringsAsFactors = FALSE
  ))
  master_xfrm <- fortify_master_xfrm(master_xfrm)

  slide_key_id <- paste0(x$master_file, x$type)
  master_key_id <- paste0(master_xfrm$file, master_xfrm$type)

  slide_xfrm_no_match <- x[!slide_key_id %in% master_key_id, ]
  slide_xfrm_no_match <- merge(slide_xfrm_no_match,
    master_ref,
    by.x = "master_file", by.y = "file",
    all.x = TRUE, all.y = FALSE
  )

  x <- merge(x, master_xfrm,
    by.x = c("master_file", "type"),
    by.y = c("file", "type"),
    all = FALSE
  )
  x$offx <- ifelse(!is.finite(x$offx), x$offx_ref, x$offx)
  x$offy <- ifelse(!is.finite(x$offy), x$offy_ref, x$offy)
  x$cx <- ifelse(!is.finite(x$cx), x$cx_ref, x$cx)
  x$cy <- ifelse(!is.finite(x$cy), x$cy_ref, x$cy)
  x$offx_ref <- NULL
  x$offy_ref <- NULL
  x$cx_ref <- NULL
  x$cy_ref <- NULL
  x$fld_id_ref <- NULL
  x$fld_type_ref <- NULL

  x <- rbind(x, slide_xfrm_no_match, stringsAsFactors = FALSE)

  i_master <- get_file_index(x$master_file)
  i_layout <- get_file_index(x$file)
  x <- x[order(i_master, i_layout, x$offy, x$offx), , drop = FALSE] # intuitive sorting: top -> bottom, left -> right
  x <- x[!(is.na(x$offx) | is.na(x$offy) | is.na(x$cx) | is.na(x$cy)),  ]

  x$type_idx <- stats::ave(x$type, x$master_file, x$file, x$type, FUN = seq_along)
  x$type_idx <- as.numeric(x$type_idx) # NB: ave returns character

  x$id <- as.integer(x$id)

  rownames(x) <- NULL # prevent meaningless rownames
  x
}


read_theme_colors <- function(doc, theme){

  nodes <- xml_find_all(doc, "//a:clrScheme/*")

  names_ <- xml_name(nodes)
  col_types_ <- xml_name(xml_children(nodes) )
  vals <- xml_attr(xml_children(nodes), "val")
  last_colors_ <- xml_attr(xml_children(nodes), "lastClr")
  vals <- ifelse(col_types_ == "srgbClr", paste0("#", vals), paste0("#", last_colors_) )
  data.frame(stringsAsFactors = FALSE, name = names_, type = col_types_, value = vals, theme = theme)
}



characterise_df <- function(x){
  names(x) <- htmlEscapeCopy(names(x))
  x <- lapply(x, function( x ) {
    if( is.character(x) ) x
    else if( is.factor(x) ) as.character(x)
    else gsub("(^ | $)+", "", format(x))
  })
  data.frame(x, stringsAsFactors = FALSE, check.names = FALSE)
}


xpath_content_selector <- "*[self::p:cxnSp or self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]"

as_xpath_content_sel <- function(prefix){
  paste0(prefix, xpath_content_selector)
}


between <- function(x, left, right ){
  x >= left & x <= right
}



simple_lag <- function( x, default=0 ){
  c(default, x[-length(x)])
}

rbind_match_columns <- function(list_df) {

  col <- unique(unlist(lapply(list_df, colnames)))
  x <- Filter(function(x) nrow(x)>0, list_df)
  x <- lapply(x, function(x, col) {
    x[, setdiff(col, colnames(x))] <- NA
    x
  }, col = col)
  do.call(rbind, x)
}

set_row_span <- function( row_details ){
  row_details$first[!row_details$first & !row_details$row_merge] <- TRUE
  row_details$row_merge <- NULL

  row_details <- split(row_details, row_details$cell_id)

  row_details <- mapply(function(dat){
    rowspan_values_at_breaks <- rle(cumsum(dat$first))$lengths
    rowspan_pos_at_breaks <- which(dat$first)
    dat$row_span <- 0L
    dat$row_span[rowspan_pos_at_breaks] <- rowspan_values_at_breaks
    dat
  }, row_details, SIMPLIFY = FALSE)
  row_details <- rbind_match_columns(row_details)
  row_details$first <- NULL
  row_details
}


#' @importFrom grDevices col2rgb rgb
is.color = function(x) {
  # http://stackoverflow.com/a/13290832/3315962
  out = sapply(x, function( x ) {
    tryCatch( is.matrix( col2rgb( x ) ), error = function( e ) F )
  })

  nout <- names(out)
  if( !is.null(nout) && any( is.na( nout ) ) )
    out[is.na( nout )] = FALSE

  out
}

correct_id <- function(doc, int_id){
  all_uid <- xml_find_all(doc, "//*[@id]")
  for(z in seq_along(all_uid) ){
    if(!grepl("[^0-9]", xml_attr(all_uid[[z]], "id"))){
      xml_attr(all_uid[[z]], "id") <- int_id
      int_id <- int_id + 1
    }
  }
  int_id
}


check_bookmark_id <- function(bkm){
  if(!is.null(bkm)){
    invalid_bkm <- is.character(bkm) &&
      length(bkm) == 1 &&
      nchar(bkm) > 0 &&
      grepl("[^:[:alnum:]_-]+", bkm)
    if(invalid_bkm){
      stop("bkm [", bkm, "] should only contain alphanumeric characters, ':', '-' and '_'.", call. = FALSE)
    }
  }
  bkm
}

is_windows <- function() {
  "windows" %in% .Platform$OS.type
}

is_doc_open <- function(file) {
  # The function checks if the `file` is open (a.k.a. is being edited).
  # This function is valid on Windows operating system only.
  suppressWarnings(file.exists(file) && !file.rename(from = file, to = file))
}


# Extract trailing numeric index in .xml filename
#
# Useful to for slideMaster and slideLayout .xml files.
#
# Examples:
#   files <- c("slideLayout1.xml", "slideLayout2.xml", "slideLayout10.xml")
#   get_file_index(files)
#
get_file_index <- function(file) {
  x <- sub(pattern = ".+?(\\d+).xml$", replacement = "\\1", x = basename(file), ignore.case = TRUE)
  as.numeric(x)
}


# Sort xml filenames by trailing numeric index
#
# Useful to for slideMaster and slideLayout xml files.
#
# Examples:
#   files <- c("slideLayout1.xml", "slideLayout2.xml", "slideLayout12.xml")
#   sort_vec_by_index(files)  # => order corresponding to trailing index
#   sort(files)           # => incorrect lexicographical ordering
#
sort_vec_by_index <- function(x) {
  indexes <- get_file_index(x)
  x[order(indexes)]
}


# Sort dataframe column by trailing index
#
# df: A dataframe
# ...: columsn to sort by, comma separated
#
# Examples:
# df <- data.frame(
#   a = paste0("file_", rep(3:1, each = 2), ".xml"),
#   b = paste0("file_", rep(3:1, 2), ".xml")
# )
# sort_dataframe_by_index(df, "a", "b")
# sort_dataframe_by_index(df, "b", "a")
#
sort_dataframe_by_index <- function(df, ...) {
  sort_columns <- c(...)
  l <- lapply(sort_columns, function(.col) {
    get_file_index(df[[.col]])
  })
  df[do.call(order, l), , drop =FALSE]
}


# rename dataframe columns
#
# Examples:
#   df_rename(mtcars, c("mpg", "cyl"), c("A", "B"))
#
df_rename <- function(df, old, new) {
  .nms <- names(df)
  .nms[match(old, .nms)] <- new
  stats::setNames(df, .nms)
}


# replacement for stopifnot() with nicer user feedback
stop_if_not_class <- function(x, class, arg = NULL) {
  check <- inherits(x, what = class)
  if (!check) {
    msg_arg <- ifelse(is.null(arg), "Incorrect input.", "Incorrect input for {.arg {arg}}")
    cli::cli_abort(c(
      msg_arg,
      "x" = "Expected {.cls {class}} but got {.cls {class(x)[1]}}"
    ), call = NULL)
  }
}


stop_if_not_rpptx <- function(x, arg = NULL) {
  stop_if_not_class(x, "rpptx", arg)
}


stop_if_not_integerish <- function(x, arg = NULL) {
  check <- is_integerish(x)
  if (!check) {
    msg_arg <- ifelse(is.null(arg), "Incorrect input.", "Incorrect input for {.arg {arg}}")
    cli::cli_abort(c(
      msg_arg,
      "x" = "Expected integerish values but got {.cls {class(x)[1]}}"
    ), call = NULL)
  }
}


#' Ensure valid slide indexes
#'
#' @param x An `rpptx` object.
#' @param idx Slide indexes.
#' @param arg Name of argument to use in error message (optional).
#' @param call Environment to display in error message. Defaults to caller env.
#'   Set `NULL` to suppress (see [cli::cli_abort]).
#' @keywords internal
stop_if_not_in_slide_range <- function(x, idx, arg = NULL, call = parent.frame()) {
  stop_if_not_rpptx(x)
  stop_if_not_integerish(idx)

  n_slides <- length(x)
  idx_available <- seq_len(n_slides)
  idx_outside <- setdiff(idx, idx_available)
  n_outside <- length(idx_outside)

  if (n_outside == 0) {
    return(invisible(NULL))
  }
  argname <- ifelse(is.null(arg), "", "of {.arg {arg}} ")
  part_1 <- paste0("{n_outside} index{?es} ", argname, "outside slide range: {.val {idx_outside}}")
  part_2 <- ifelse(n_slides == 0,
    "Presentation has no slides!",
    "Slide indexes must be in the range [{min(idx_available)}..{max(idx_available)}]"
  )
  cli::cli_abort(c(part_1, "x" = part_2), call = call)
}


check_unit <- function(unit, choices, several.ok = FALSE) {
  if (!several.ok && length(unit) != 1) {
    cli::cli_abort(
      c("{.arg unit} is not length 1.",
        "x" = "{.arg unit} must be {.emph a string}."
      )
    )
  }
  if (!unit %in% choices) {
    cli::cli_abort(
      c("{.arg unit} should be one of {.or {choices}}.",
        "x" = "{.arg unit} was {.emph {unit}\"}."
      )
    )
  }

  unit
}


#' Update a vector or list of named defaults values
#'
#' Helper to update named default values. Takes care of most common cases and
#' returns sensible error messages.
#'
#' @param x A (named) vector or list to update the default values with.
#' @param default A *named* vector or list with default values.
#' @param argname Name of the arg. Used in error message, defaults to `"x"`.
#' @param default_if_null Returns defaults values if `x` is `NULL`? (default is `TRUE`)
#' @param partial Enable partial name matching= (default `TRUE`).
#' @param as_list Return a list (default `TRUE`) or a named vector?
#' @return List of update default values.
#' @noRd
#' @examples
#' defaults <- c(aa = 1, bb = 2, cc = 3)
#' update_named_defaults(c(a = 99, b = 0), defaults) # partial name match
#' update_named_defaults(c(a = 99, b = 0), defaults, as_list = FALSE) # return as vector
#' update_named_defaults(3:1, defaults) # match by position
#' update_named_defaults(2, defaults)
#' update_named_defaults(NULL, defaults, default_if_null = TRUE)
#'
update_named_defaults <- function(x, default, argname = "x", default_if_null = TRUE, partial = TRUE, as_list = TRUE) {

  if (default_if_null && is.null(x)) {
    res <- as.list(default)
    if (!as_list) res <- unlist(res)
    return(res)
  }

  x <- as.list(x)
  default <- as.list(default)

  if (!is_named(default)) {
    cli::cli_abort(
      c("Some default vector elements have no names",
        "x" = "{.arg default} must be a named vector"
      ), call = NULL
    )
  }
  x <- as.list(x)
  len_x <- length(x)
  len_default <- length(default)
  names_default <- names(default)
  if (len_x > len_default) {
    cli::cli_abort(c(
      "Length of {.arg {argname}} ({.val {len_x}}) exceeds length of {.arg default} ({.val {len_default}})",
      "x" = "Length of {.arg x} must be smaller or equal to the length of {.arg default}"
    ), call = NULL)
  }

  # unnamed case => convert to named case
  if (!is_named(x)) {
    if (len_x == 1) {
      x <- rep(x, len_default)
      len_x <- length(x)
    }
    if (len_x != len_default) {
      cli::cli_abort(c(
        "{.arg {argname}} has incorrect length ({len_x})",
        "x" = "If {.arg {argname}} has no names, it must be length 1 or the length of {.arg default} ({len_default})"
      ), call = NULL)
    }
    names(x) <- names_default
  }

  # named case
  if (partial) { # => partial name matching
    matched <- pmatch(names(x), names_default, duplicates.ok = TRUE)
  } else { # exact name matching
    matched <- match(names(x), names_default, nomatch = NA)
  }
  nms_new <- ifelse(is.na(matched), NA, names_default[matched])
  i_na <- is.na(nms_new)

  if (any(i_na)) {
    msg_partial <- ifelse(partial, "Partial name matching is supported", "Partial name matching is not enabled")
    cli::cli_abort(
      c("Found {sum(i_na)} unknown name{?s} in {.arg {argname}}: {.val {names(x)[i_na]}}",
        "x" = "{.arg {argname}} understands {.val {names_default}}",
        "i" = cli::col_silver(msg_partial)
      ),
      call = NULL
    )
  }
  # duplicate position
  ii_dupes <- duplicated(nms_new)
  if (any(ii_dupes)) {
    cli::cli_abort(
      c("Duplicate entries in {.arg location}: {.val {unique(nms_new[ii_dupes])}}",
        "x" = "Each name in {.arg location} must be unique",
        "i" = cli::col_silver("Partial name matching is supported")
      ),
      call = NULL
    )
  }
  x <- setNames(x, nms_new)
  res <- utils::modifyList(x = default, val = as.list(x), keep.null = TRUE)
  if (!as_list) res <- unlist(res)
  res
}


# htmlEscapeCopy ----

htmlEscapeCopy <- local({

  .htmlSpecials <- list(
    `&` = '&amp;',
    `<` = '&lt;',
    `>` = '&gt;'
  )
  .htmlSpecialsPattern <- paste(names(.htmlSpecials), collapse='|')
  .htmlSpecialsAttrib <- c(
    .htmlSpecials,
    `'` = '&#39;',
    `"` = '&quot;',
    `\r` = '&#13;',
    `\n` = '&#10;'
  )
  .htmlSpecialsPatternAttrib <- paste(names(.htmlSpecialsAttrib), collapse='|')
  function(text, attribute=FALSE) {
    pattern <- if(attribute)
      .htmlSpecialsPatternAttrib
    else
      .htmlSpecialsPattern
    text <- enc2utf8(as.character(text))
    # Short circuit in the common case that there's nothing to escape
    if (!any(grepl(pattern, text, useBytes = TRUE)))
      return(text)
    specials <- if(attribute)
      .htmlSpecialsAttrib
    else
      .htmlSpecials
    for (chr in names(specials)) {
      text <- gsub(chr, specials[[chr]], text, fixed = TRUE, useBytes = TRUE)
    }
    Encoding(text) <- "UTF-8"
    return(text)
  }
})


# metric units -----------------------------------------------
#
cm_to_inches <- function(x) {
  x / 2.54
}

mm_to_inches <- function(x) {
  x / 25.4
}

convin <- function(unit, x) {
  unit <- match.arg(unit, choices = c("in", "cm", "mm"), several.ok = FALSE)
  if (!identical("in", unit)) {
    x <- do.call(paste0(unit, "_to_inches"), list(x = x))
  }
  x
}


# from rlang pkg  ------------------------------------------

is_named <- function(x) {
  nms <- names(x)
  if (is.null(nms)) {
    return(FALSE)
  }
  if (any(detect_void_name(nms))) {
    return(FALSE)
  }
  TRUE
}


detect_void_name <- function(x) {
  x == "" | is.na(x)
}


# is_integerish(1)
# is_integerish(1.0)
# is_integerish(c(1.0, 2.0))
is_integerish <- function(x) {
  ii <- all(is.numeric(x) | is.integer(x))
  jj <- all(x == as.integer(x))
  ii && jj
}


# coalesce for R
`%||%` <- function(l, r) {
  if (is.null(l)) r else l
}


# file ops  ------------------------------------------

#' Opens a file locally
#'
#' Opening a file locally requires a compatible application to be installed
#' (e.g., MS Office or LibreOffice for .pptx or .docx files).
#'
#' @details
#' *NB:* Function is a small wrapper around [utils::browseURL()] to have a more
#' suitable function name.
#'
#' @param path File path.
#' @export
#' @examples
#' x <- read_pptx()
#' x <- add_slide(x, "Title Slide", ctrTitle = "My Title")
#' file <- print(x, tempfile(fileext = ".pptx"))
#' \dontrun{
#' open_file(file)}
open_file <- function(path) {
  path <- normalizePath(path, winslash = "/", mustWork = TRUE)
  utils::browseURL(path)
}
