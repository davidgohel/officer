#' @export
#' @title generic container
#' @format a dummy generic container
#' @field new(node) create a container, if \code{node} is missing an empty container is returned.
#' @field set_content(text) set content, \code{node} is containing xml code in a single string.
#' @field format() get xml code
#' @field print() print object
#' @importFrom R6 R6Class
any_content <- R6Class(
  "any_content",
  public = list(
    content = NULL,

    initialize = function(node = NULL) {
      if( !is.null(node) && !inherits(node, "xml_missing") )
        self$content <- as.character(node)
    },
    set_content = function( text ){
      self$content <- text
      self
    },
    format = function() {
      ifelse( is.null(self$content), "", self$content )
    },
    print = function() {
      print(self$content)
    }
  )
)

