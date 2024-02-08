#' @title Get comments in a Word document as a data.frame
#' @description return a data.frame representing the comments in a Word document.
#' @param x an rdocx object
#' @details
#' Each row of the returned data frame contains data for one comment. The
#'    columns contain the following information:
#' * "comment_id" - unique comment id
#' * "author" - name of the comment author
#' * "initials" - initials of the comment author
#' * "date" - timestamp of the comment
#' * "text" - a list column of characters containing the comment text. Elements can
#'    be vectors of length > 1 if a comment contains multiple paragraphs,
#'    blocks or runs or of length 0 if the comment is empty.
#' * "para_id" - a list column of characters containing the parent paragraph IDs.
#'    Elememts can be vectors of length > 1 if a comment spans multiple paragraphs
#'    or of length 0 if the comment has no parent paragraph.
#' * "commented_text" - a list column of characters containing the
#'    commented text. Elememts can be vectors of length > 1 if a comment
#'    spans multiple paragraphs or runs or of length 0 if the commented text is empty.
#' @examples
#' bl <- block_list(
#'   fpar("Comment multiple words."),
#'   fpar("Second line")
#' )
#'
#' a_par <- fpar(
#'   "This paragraph contains",
#'   run_comment(
#'     cmt = bl,
#'     run = ftext("a comment."),
#'     author = "Author Me",
#'     date = "2023-06-01"
#'   )
#' )
#'
#' doc <- read_docx()
#' doc <- body_add_fpar(doc, value = a_par, style = "Normal")
#'
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#'
#' docx_comments(read_docx(docx_file))
#' @export
docx_comments <- function(x) {
  stopifnot(inherits(x, "rdocx"))

  comment_ids <- xml_attr(
    xml_find_all(
      x$doc_obj$get(), "/w:document/w:body//*[self::w:commentRangeStart]"
    ), "id"
  )

  comment_text_runs <- lapply(comment_ids, function(id) {
    xml_find_all(
      x$doc_obj$get(),
      paste0(
        "/w:document/w:body//*[self::w:r[w:t and",
        "preceding::w:commentRangeStart[@w:id=\'", id, "\']",
        " and ",
        "following::w:commentRangeEnd[@w:id=\'", id, "\']]]"
      )
    )
  })

  data <- data.frame(
    comment_id = comment_ids
  )
  # Add parent paragraph id
  data$para_id <- lapply(
    comment_text_runs,
    function(x) xml_attr(xml_parent(x), "paraId")
  )
  data$commented_text <- lapply(comment_text_runs, xml_text)

  comments <- xml_find_all(x$comments$get(), "//w:comments/w:comment")

  out <- data.frame(
    stringsAsFactors = FALSE,
    comment_id = xml_attr(comments, "id"),
    author = xml_attr(comments, "author"),
    initials = xml_attr(comments, "initials"),
    date = xml_attr(comments, "date")
  )

  out$text <- lapply(
    comments,
    function(x) xml_text(xml_find_all(x, "w:p/w:r/w:t"))
  )

  data <- merge(out, data, by = "comment_id", all.x = TRUE)
  data[order(as.integer(data$comment_id)), ]
}
