#' @export
#' @title Empty block for 'PowerPoint'
#' @description Create an empty object to include as an empty placeholder shape in a
#' presentation. This comes in handy when presentation are updated
#' through R, but a user still wants to add some comments in this new content.
#'
#' Empty content also works with layout fields (slide number and date) to preserve them:
#' they are included on the slide and keep being updated by PowerPoint, i.e. update to the
#' when the slide number when the slide moves in the deck, update to the date.
#'
#' @examples
#' fileout <- tempfile(fileext = ".pptx")
#' doc <- read_pptx()
#' doc <- add_slide(doc, layout = "Two Content",
#'   master = "Office Theme")
#' doc <- ph_with(x = doc, value = empty_content(),
#'  location = ph_location_type(type = "title") )
#'
#' doc <- add_slide(doc)
#' # add slide number as a computer field
#' doc <- ph_with(
#'   x = doc, value = empty_content(),
#'   location = ph_location_type(type = "sldNum"))
#'
#' print(doc, target = fileout )
#' @seealso [ph_with()], [body_add_blocks()]
empty_content <- function(){
  x <- list()
  class(x) <- "empty_content"
  x
}
