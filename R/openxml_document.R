# openxml_document --------------------------------------------------------
# This class handle basic operations on a openxml document:
# - initialize
openxml_document <- R6Class(
  "openxml_document",
  public = list(

    initialize = function( dir ) {
      private$reldir = dir
      private$rels_doc <- relationship$new()
    },

    feed = function( file ) {
      private$filename <- file
      private$doc <- read_xml(file)

      private$rels_filename <- file.path( dirname(file), "_rels", paste0(basename(file), ".rels") )

      if( file.exists(private$rels_filename) )
        private$rels_doc <- relationship$new()$feed_from_xml(private$rels_filename)
      else private$rels_doc <- relationship$new()

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
    replace_xml = function(file){
      # dont use `options = "NSCLEAN"` otherwise invalid document as a result
      private$doc <- read_xml(file, options = "NOBLANKS")
    },
    dir_name = function(){
      private$reldir
    },
    save = function() {
      # remove duplicate namespace definitions
      private$doc <- read_xml(as.character(private$doc), options = "NSCLEAN")
      write_xml(private$doc, file = private$filename)
      if( nrow(self$rel_df()) > 0 ){
        private$rels_doc$write(private$rels_filename)
      }
      self
    },
    remove = function() {
      unlink(private$filename)
      if( file.exists(private$rels_filename) )
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
