docx_part <- R6Class(
  "docx_part",
  inherit = openxml_document,
  public = list(

    initialize = function( path, main_file, cursor, body_xpath ) {
      super$initialize("word")
      private$package_dir <- path
      private$body_xpath <- body_xpath
      super$feed(file.path(private$package_dir, "word", main_file))
      private$cursor <- cursor
    },
    length = function( ){
      xml_length( xml_find_first(self$get(), private$body_xpath ) )
    }

  ),
  private = list(
    package_dir = NULL,
    cursor = NULL,
    body_xpath = NULL
  )

)
