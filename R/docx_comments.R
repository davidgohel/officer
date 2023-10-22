#' @title Get comments in a Word document as a data.frame
#' @description return a data.frame representing the comments in a Word document.
#' @param x an rdocx object
#' @examples
#' bl <- block_list(
#'   fpar("Comment multiple words."),
#'   fpar("Second line")
#' )
#'
#' a_par <- fpar(
#'   "This paragraph contains",
#'   run_comment(
#'   cmt = bl,
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

  comment_nodes <- xml_find_all(
    x$doc_obj$get(), "//*[self::w:p/w:commentRangeStart]"
  )

  if (length(comment_nodes) > 0) {
    data <- lapply(comment_nodes, comment_as_tibble)
    data <- rbind_match_columns(data)
  } else {
    data <- data.frame(
      comment_id = integer(0),
      commented_text = character(0)
    )
  }

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

  data
}

comment_as_tibble <- function(node) {
  node_name <- xml_name(node)
  name_children <- xml_name(xml_children(node))

  comment_range <- grep("commentRange", name_children)

  comment_data <- data.frame(
    comment_id = xml_attr(xml_child(node, comment_range[[1]]), "id"),
    stringsAsFactors = FALSE
  )
  comment_range <- seq(comment_range[[1]] + 1, comment_range[[2]] - 1)
  comment_data$commented_text <-
    paste0(
      vapply(
        comment_range,
        function(x) xml_text(xml_child(node, x)),
        character(1)
      ),
      collapse = ""
    )

  comment_data
}
