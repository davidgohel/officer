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
    },
    clean_media = function() {
      # get package directory
      package_dir <- dirname(dirname(dirname(private$rels_filename)))
      # find .xml.rels files
      rel_files <- list.files(package_dir, pattern = "\\.xml.rels$", recursive = TRUE, full.names = TRUE)

      # collect media files that are referenced in .xml.rels files
      media_file_rels <- lapply(rel_files, function(xml){
        zz <- xml2::read_xml(xml)
        rels <- xml2::xml_children(zz)
        targets <- xml2::xml_attr(rels, "Target")
        grep("^\\.\\./media/", targets, value = TRUE)
      })
      media_file_rels <- file.path(dirname(dirname(private$rels_filename)), sub("\\.\\./", "", unique(unlist(media_file_rels))))

      # get files in media folder
      media_folder <- grep("/media$", list.dirs(package_dir), value = TRUE)
      media_files <- list.files(media_folder, recursive = TRUE, full.names = TRUE)

      # delete unused files
      files_to_delete <- media_files[!(media_files %in% media_file_rels)]
      unlink(files_to_delete, force = TRUE)
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
