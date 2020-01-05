#' @export
#' @title create empty blocks
#' @description an empty object to include as an empty placeholder shape in a
#' presentation. This comes in handy when presentation are updated
#' through R, but a user still wants to write the takeaway statements in
#' PowerPoint.
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Two Content",
#'   master = "Office Theme")
#' doc <- ph_with(x = doc, value = empty_content(),
#'  location = ph_location_type(type = "title") )
#' print(doc, target = fileout )
#' @seealso [ph_with()], [body_add_blocks()]
empty_content <- function(){
  x <- list()
  class(x) <- "empty_content"
  x
}
