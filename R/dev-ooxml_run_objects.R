# comment ----

#' @export
#' @title Comment for 'Word'
#' @description Add a comment on a run object.
#' @param cmt comment to associate with run.
#' @param run a run object, made with a call to one of
#' the "run functions for reporting".
#' @inheritSection ftext usage
#' @examples
#' ft <- fp_text(font.size = 12, bold = TRUE)
#' run_bookmark("par1", ftext("some text", ft))
#' @family run functions for reporting
run_comment <- function(x, run) {

  all_run <- FALSE

  if(inherits(run, "run")){
    run <- list(run)
  }

  if(is.list(run) && !inherits(run, "run")){
    all_run <- all(vapply(run, inherits, FUN.VALUE = FALSE, what = "run" ))
  }


  if(!all_run)
    stop("`run` must be a run object (ftext for example) or a list of run objects.")

  z <- list(comment = x, run = run)
  class(z) <- c("run_comment", "run")
  z
}

#' @export
to_wml.run_comment <- function(x, add_ns = FALSE, ...) {
  open_tag <- wr_ns_no
  if (add_ns) {
    open_tag <- wr_ns_yes
  }

  x$footnote[[1]]$chunks <- append(list(run_footnoteref(x$pr)), x$footnote[[1]]$chunks)
  blocks <- sapply(x$footnote, to_wml)
  blocks <- paste(blocks, collapse = "")

  id <- basename(tempfile(pattern = "footnote"))
  base_ns <- "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
  footnote_xml <- paste0(
    "<w:footnote ", base_ns, " w:id=\"",
    id,
    "\">",
    blocks, "</w:footnote>"
  )

  footnote_ref_xml <- paste0(
    open_tag,
    if (!is.null(x$pr)) rpr_wml(x$pr),
    "<w:commentReference w:id=\"",
    id,
    "\">",
    footnote_xml,
    "</w:commentReference>",
    "</w:r>"
  )

  footnote_ref_xml
}

#' @export
to_wml.run_comment <- function(x, add_ns = FALSE, ...) {
  runs <- lapply(x$run, to_wml, add_ns = add_ns, ...)
  runs <- do.call(paste0, runs)
  comment(str = runs)
}

# bookmark -----

comment <- function(str) {
  new_id <- uuid_generate()
  cmt_start_str <- sprintf("<w:commentRangeStart w:id=\"%s\"/>", new_id)
  cmt_start_end <- sprintf("<w:commentRangeEnd w:id=\"%s\"/>", new_id)
  cmt_ref <- sprintf("<w:r><w:commentReference w:id=\"%s\"/></w:r>", new_id)

  paste0(cmt_start_str, str, cmt_ref, cmt_start_end)
}
