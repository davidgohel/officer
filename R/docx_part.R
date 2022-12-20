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
    get_at_cursor = function() {
      # prevent some packages that use internals to fail on cran
      # to developpers: you should not use internals, it makes debugging/fixing a real pain
      xml_elt <- paste0(wp_ns_yes,"<w:r><w:t></w:t></w:r></w:p>")
      as_xml_document(xml_elt)
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
