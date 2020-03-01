#' @importFrom R6 R6Class
relationship <- R6Class(
  "relationship",
  public = list(

    initialize = function(id = character(0), type = character(0), target = character(0)) {
      private$id <- id
      private$type <- type
      private$target <- target
      private$target_mode <- as.character( rep(NA, length(target)) )
      private$ext_src <- character(length(id))
    },
    feed_from_xml = function(path) {
      doc <- read_xml( x = path )
      children <- xml_children( doc )
      ns <- xml_ns( doc )

      private$id <- c( private$id, sapply(children, xml_attr, attr = "Id", ns) )
      private$type <- c( private$type, sapply( children, xml_attr, attr = "Type", ns) )
      private$target <- c( private$target, sapply( children, xml_attr, attr = "Target", ns) )
      private$target_mode <- c( private$target_mode, sapply( children, xml_attr, attr = "TargetMode", ns) )
      private$ext_src <- c( private$ext_src, character(length(children)) )
      self
    },
    write = function(path) {
      if( length(private$id) )
        str <- paste0("<Relationship Id=\"", private$id,
                    "\" Type=\"", private$type,
                    "\" Target=\"", private$target,
                    ifelse(
                          is.na(private$target_mode),
                          "",
                          "\" TargetMode=\"External" ),
                    "\"/>", collapse = "")
      else str <- ""
      str <- paste0("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
             "\n<Relationships  xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">", str, "</Relationships>")
      dir.create(dirname(path), showWarnings = FALSE, recursive = FALSE)
      cat(str, file = path)
      self
    },
    get_next_id = function(){
      max(c(0, private$get_int_id() ), na.rm = TRUE ) + 1
    },
    get_data = function() {
      if( length(private$id) ){
        data <- data.frame(id = private$id,
                           int_id = as.integer(gsub("rId([0-9]+)", "\\1", private$id)),
                           type = private$type,
                           target = private$target,
                           target_mode = private$target_mode,
                           ext_src = private$ext_src,
                           stringsAsFactors = FALSE )
        data[order(data$id),]
      } else {
        data <- data.frame(id = character(0),
                        int_id = integer(0),
                        type = character(0),
                        target = character(0),
                        target_mode = character(0),
                        ext_src = character(0),
                        stringsAsFactors = FALSE )
      }
      data
    },
    get_images_path = function() {
      is_img <- basename( private$type ) %in% "image"
      targets <- private$target[is_img]
      names( targets ) <- private$id[is_img]
      targets
    },
    add_img = function( src, root_target ) {

      src <- setdiff(src, private$ext_src)
      if( !length(src) ) return(self)

      if( any( grepl(" ", basename(src) ) ) ){
        stop(paste(src, collapse = ", "),
             ": images with blanks in their basenames are not supported, please rename the file(s).",
             call. = FALSE)
      }

      last_id <- max( c(0, private$get_int_id() ), na.rm = TRUE )

      id <- paste0("rId", seq_along(src) + last_id)
      type <- rep("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                 length(src))
      target <- file.path(root_target, basename(src) )

      private$id <- c( private$id, id )
      private$type <- c( private$type, type )
      private$target <- c( private$target, target )
      private$target_mode <- c( private$target_mode, rep(NA, length(id) ) )
      private$ext_src <- c( private$ext_src, src )

      self
    },
    add_drawing = function( src, root_target ) {

      src <- setdiff(src, private$ext_src)
      if( !length(src) ) return(self)
      last_id <- max( c(0, private$get_int_id()), na.rm = TRUE )

      id <- paste0("rId", seq_along(src) + last_id)
      type <- rep("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                 length(src))
      target <- file.path(root_target, basename(src) )

      private$id <- unlist( c( private$id, id ) )
      private$type <- unlist( c( private$type, type ) )
      private$target <- unlist( c( private$target, target ) )
      private$target_mode <- unlist( c( private$target_mode, rep(NA, length(id) ) ) )
      private$ext_src <- unlist( c( private$ext_src, rep(NA, length(id) ) ) )

      self
    },
    add = function(id, type, target, target_mode = NA ) {
      if( !target %in% private$target ){
        private$id <- c( private$id, id )
        private$type <- c( private$type, type )
        private$target <- c( private$target, target )
        private$target_mode <- c( private$target_mode, target_mode )
        private$ext_src <- c( private$ext_src, "" )
      }
      self
    },
    remove = function( target ) {
      id <- which( basename(private$target) %in% basename(target)  )
      private$id <- private$id[-id]
      private$type <- private$type[-id]
      private$target <- private$target[-id]
      private$target_mode <- private$target_mode[-id]
      private$ext_src <- private$ext_src[-id]
      self
    },
    show = function() {
      print(self$get_data())
    }
  ),
  private = list(
    id = NA,
    type = NA,
    target = NA,
    target_mode = NA,
    ext_src = NA,
    get_int_id = function(){
      if( length(private$id) > 0 )
        as.integer(gsub("rId([0-9]+)", "\\1", private$id))
      else integer(0)
    }
  )
)

