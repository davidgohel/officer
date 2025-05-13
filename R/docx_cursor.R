#' @export
#' @rdname cursor
#' @title Set cursor in a 'Word' document
#' @description A set of functions is available to manipulate
#' the position of a virtual cursor. This cursor will be used when
#' inserting, deleting or updating elements in the document.
#' @section cursor_begin:
#' Set the cursor at the beginning of the document, on the first element
#' of the document (usually a paragraph or a table).
#' @param x a docx device
#' @examples
#' library(officer)
#'
#' # create a template ----
#' doc <- read_docx()
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "Hello text to replace")
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "blah blah blah")
#' doc <- body_add_par(doc, "Hello text to replace")
#' doc <- body_add_par(doc, "blah blah blah")
#' template_file <- print(
#'   x = doc,
#'   target = tempfile(fileext = ".docx")
#' )
#'
#' # replace all pars containing "to replace" ----
#' doc <- read_docx(path = template_file)
#' while (cursor_reach_test(doc, "to replace")) {
#'   doc <- cursor_reach(doc, "to replace")
#'
#'   doc <- body_add_fpar(
#'     x = doc,
#'     pos = "on",
#'     value = fpar(
#'       "Here is a link: ",
#'       hyperlink_ftext(
#'         text = "yopyop",
#'         href = "https://cran.r-project.org/"
#'       )
#'     )
#'   )
#' }
#'
#' doc <- cursor_end(doc)
#' doc <- body_add_par(doc, "Yap yap yap yap...")
#'
#' result_file <- print(
#'   x = doc,
#'   target = tempfile(fileext = ".docx")
#' )
cursor_begin <- function(x) {
  if (length(x$officer_cursor$nodes_names) > 0L) {
    x$officer_cursor$which <- 1L
  } else {
    x$officer_cursor$which <- 0L
  }
  x
}

#' @rdname cursor
#' @param id bookmark id
#' @section cursor_bookmark:
#' Set the cursor at a bookmark that has previously been set.
#' @export
#' @examples
#'
#' # cursor_bookmark ----
#'
#' doc <- read_docx()
#' doc <- body_add_par(doc, "centered text", style = "centered")
#' doc <- body_bookmark(doc, "text_to_replace")
#' doc <- body_add_par(doc, "A title", style = "heading 1")
#' doc <- body_add_par(doc, "Hello world!", style = "Normal")
#' doc <- cursor_bookmark(doc, "text_to_replace")
#' doc <- body_add_table(doc, value = iris, style = "table_template")
#'
#' print(doc, target = tempfile(fileext = ".docx"))
cursor_bookmark <- function(x, id) {
  xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", id)
  bm_start <- xml_find_first(x$doc_obj$get(), xpath_)

  if (inherits(bm_start, "xml_missing")) {
    stop("cannot find bookmark ", shQuote(id), call. = FALSE)
  }

  bm_id <- xml_attr(bm_start, "id")

  nodes_with_text <- xml_find_all(
    x$doc_obj$get(),
    "/w:document/w:body/*|/w:ftr/*|/w:hdr/*"
  )
  test_start <- sapply(nodes_with_text, function(node) {
    expr <- sprintf("/descendant::w:bookmarkStart[@w:id='%s']", bm_id)
    match_node <- xml_child(node, expr)
    !inherits(match_node, "xml_missing")
  })
  if (!any(test_start)) {
    stop("bookmark ", shQuote(id), " has not been found in the document", call. = FALSE)
  }

  test_end <- sapply(nodes_with_text, function(node) {
    expr <- sprintf("/descendant::w:bookmarkEnd[@w:id='%s']", bm_id)
    match_node <- xml_child(node, expr)
    !inherits(match_node, "xml_missing")
  })

  on_same_par <- test_start == test_end
  if (!all(on_same_par)) {
    stop("bookmark ", shQuote(id), " does not end in the same paragraph (or is on the whole paragraph)", call. = FALSE)
  }

  x$officer_cursor$which <- which(test_start)[1]

  x
}

#' @export
#' @rdname cursor
#' @section cursor_end:
#' Set the cursor at the end of the document, on the last element
#' of the document.
cursor_end <- function(x) {
  x$officer_cursor$which <- length(x$officer_cursor$nodes_names)
  x
}

#' @export
#' @rdname cursor
#' @param keyword keyword to look for as a regular expression
#' @param fixed logical. If TRUE, pattern is a string to be matched as is.
#' @section cursor_reach:
#' Set the cursor on the first element of the document
#' that contains text specified in argument `keyword`.
#' The argument `keyword` is a regexpr pattern.
cursor_reach <- function(x, keyword, fixed = FALSE) {
  nodes_with_text <- xml_find_all(
    x$doc_obj$get(),
    "/w:document/w:body/*|/w:ftr/*|/w:hdr/*"
  )

  if (length(nodes_with_text) < 1) {
    stop("no text found in the document", call. = FALSE)
  }

  text_ <- xml_text(nodes_with_text)
  test_ <- grepl(pattern = keyword, x = text_, fixed = fixed)
  if (!any(test_)) {
    stop(keyword, " has not been found in the document", call. = FALSE)
  }
  x$officer_cursor$which <- which(test_)[1]
  x
}

#' @export
#' @rdname cursor
#' @param keyword keyword to look for as a regular expression
#' @section cursor_reach_test:
#' Test if an expression has a match in the document
#' that contains text specified in argument `keyword`.
#' The argument `keyword` is a regexpr pattern.
cursor_reach_test <- function(x, keyword) {
  nodes_with_text <- xml_find_all(
    x$doc_obj$get(),
    "/w:document/w:body/*|/w:ftr/*|/w:hdr/*"
  )

  if (length(nodes_with_text) < 1) {
    stop("no text found in the document", call. = FALSE)
  }

  text_ <- xml_text(nodes_with_text)
  test_ <- grepl(pattern = keyword, x = text_)
  any(test_)
}

#' @export
#' @rdname cursor
#' @section cursor_forward:
#' Move the cursor forward, it increments the cursor in the document.
cursor_forward <- function(x) {
  x$officer_cursor$which <- min(c(length(x$officer_cursor$nodes_names), x$officer_cursor$which + 1L))
  x
}

#' @export
#' @rdname cursor
#' @section cursor_backward:
#' Move the cursor backward, it decrements the cursor in the document.
cursor_backward <- function(x) {
  x$officer_cursor$which <- max(c(0L, x$officer_cursor$which - 1L))
  x
}

# officer docx cursor ----
officer_cursor <- function(node) {
  nodes <- xml_find_all(node, "/w:document/w:body/*")
  nodes_names <- xml_name(nodes)
  nodes_names <- nodes_names[!nodes_names %in% "sectPr"]
  x <- list(
    nodes_names = nodes_names,
    which = length(nodes_names)
  )
  class(x) <- "officer_cursor"
  x
}


#' @export
as.character.officer_cursor <- function(x, ...) {
  if (length(x$nodes_names) < 1) {
    return(NA_character_)
  }
  paste0("/w:document/w:body/*[", x$which, "]")
}

## xml on cursor ----
ooxml_on_cursor <- function(x, node) {
  if (length(x$nodes_names) < 1) {
    return(NULL)
  }
  node <- xml_find_first(node, as.character(x))
  if (inherits(node, "xml_missing")) {
    stop("cursor does not correspond to any node", call. = FALSE)
  }
  node
}


## cursor feed ----
cursor_append <- function(x, what) {
  x$nodes_names <- c(x$nodes_names, what)
  x$which <- x$which + 1L
  x
}
cursor_add_after <- function(x, what) {
  seq_left <- seq_along(x$which)
  set_left <- x$nodes_names[seq_left]
  set_right <- x$nodes_names[-seq_left]
  x$nodes_names <- c(set_left, what, set_right)
  x$which <- x$which + 1L
  x
}

cursor_add_before <- function(x, what) {
  seq_left <- seq_along(x$which - 1)
  set_left <- x$nodes_names[seq_left]
  set_right <- x$nodes_names[-seq_left]
  x$nodes_names <- c(set_left, what, set_right)
  x
}
cursor_replace_nodename <- function(x, what) {
  x$nodes_names[x$which] <- what
  x
}
cursor_delete <- function(x, what) {
  stopifnot(x$nodes_names[x$which] == what)
  x$nodes_names <- x$nodes_names[-x$which]
  if (x$which > length(x$nodes_names)) {
    x$which <- x$which - 1L
  }
  x
}
