# openxml_document --------------------------------------------------------
openxml_document <- R6Class(
  "openxml_document",
  public = list(

    initialize = function( dir ) {
      private$reldir = dir
    },

    feed = function( file ) {
      private$filename <- file
      private$rels_filename <- file.path( dirname(file), "_rels", paste0(basename(file), ".rels") )

      private$doc <- read_xml(file)
      private$rels_doc <- relationship$new()$feed_from_xml(private$rels_filename)
      self
    },
    file_name = function(){
      private$filename
    },
    name = function(){
      basename(private$filename)
    },
    get = function(){
      private$doc
    },
    dir_name = function(){
      private$reldir
    },
    save = function() {
      write_xml(private$doc, file = private$filename)
      private$rels_doc$write(private$rels_filename)
      self
    },
    remove = function() {
      unlink(private$filename)
      unlink(private$rels_filename)
      private$filename
    },
    rel_df = function(){
      private$rels_doc$get_data()
    },
    relationship = function(){
      private$rels_doc
    }

  ),
  private = list(

    filename = NULL,
    rels_filename = NULL,
    doc = NULL,
    rels_doc = NULL,
    reldir = NULL

  )
)
