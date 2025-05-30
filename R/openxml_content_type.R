attr_chunk <- function( x ){
  if( !is.null(x) && length( x ) > 0){
    attribs <- paste0(names(x), "=", shQuote(x, type = "cmd"), collapse = " " )
    attribs <- paste0(" ", attribs)
  } else attribs <- ""
  attribs
}

# content_type -----
content_type <- R6Class(
  "content_type",
  public = list(

    initialize = function( package_dir ) {

      private$filename <- file.path(package_dir, "[Content_Types].xml")

      doc <- read_xml(x = private$filename )
      ns <- xml_ns(doc)

      node_template <- xml_find_first(doc, "d1:Override[@ContentType='application/vnd.openxmlformats-officedocument.presentationml.template.main+xml']")
      if (!inherits(node_template, "xml_missing")) {
        xml_attr(node_template, 'ContentType') <- "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
      }
      node_template <- xml_find_first(doc, "d1:Override[@ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml']")
      if (!inherits(node_template, "xml_missing")) {
        xml_attr(node_template, 'ContentType') <- "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
      }

      extension <- xml_find_all(doc, "//*[contains(local-name(), 'Default')]/@Extension", ns = ns)
      extension <- xml_text( extension )
      content_type <- xml_find_all(doc, "//*[contains(local-name(), 'Default')]/@ContentType", ns = ns)
      content_type <- xml_text( content_type )
      names(content_type) <- extension
      private$default <- content_type

      partname <- xml_find_all(doc, "//*[contains(local-name(), 'Override')]/@PartName", ns = ns)
      partname <- xml_text(partname)
      content_type <- xml_find_all(doc, "//*[contains(local-name(), 'Override')]/@ContentType", ns = ns)
      content_type <- xml_text(content_type)
      names(content_type) <- partname
      private$override <- content_type

    },

    add_slide = function(partname){
      partname <- basename(partname)
      partname <- file.path("/ppt", "slides", partname )
      content_type <- setNames("application/vnd.openxmlformats-officedocument.presentationml.slide+xml", partname )
      override <- c( private$override, content_type )
      private$override <- override
      self
    },
    add_notesSlide = function(partname){
      partname <- basename(partname)
      partname <- file.path("/ppt", "notesSlides", partname )
      content_type <- setNames("application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", partname )
      override <- c( private$override, content_type )
      private$override <- override
      self
    },
    add_override = function(value){
      if (!names(value) %in% names(private$override)){
        override <- c( private$override, value )
        private$override <- override
      }
      self
    },
    remove_slide = function(partname){
      id <- which( basename(names(private$override)) %in% basename(partname) )
      private$override <- private$override[-id]
      self
    },
    add_ext = function( extension, type ){
      if( !type %in% private$default && !extension %in% names(private$default) ){
        content_type <- setNames(type, extension )
        default <- c( private$default, content_type )
        private$default <- default
      }
      self
    },

    save = function() {
      self$add_ext(extension = "jpeg", type = "image/jpeg")
      self$add_ext(extension = "gif", type = "image/gif")
      self$add_ext(extension = "png", type = "image/png")
      self$add_ext(extension = "svg", type = "image/svg+xml")
      self$add_ext(extension = "bmp", type = "image/bmp")
      self$add_ext(extension = "emf", type = "image/x-emf")
      self$add_ext(extension = "wmf", type = "image/x-wmf")
      self$add_ext(extension = "tiff", type = "image/tiff")
      self$add_ext(extension = "pdf", type = "application/pdf")
      self$add_ext(extension = "jpg", type = "application/octet-stream")
      attribs <- attr_chunk(c(xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"))
      out <- paste0("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                    "\n<Types", attribs, ">")
      if(length(private$default) > 0 ){
        default <- sprintf("<Default Extension=\"%s\" ContentType=\"%s\"/>", names(private$default), private$default )
        default <- paste0(default, collapse = "")
        out <- paste0(out, default )
      }
      if(length(private$override) > 0 ){
        override <- sprintf("<Override PartName=\"%s\" ContentType=\"%s\"/>", names(private$override), private$override )
        override <- paste0(override, collapse = "")
        out <- paste0(out, override )
      }
      out <- paste0(out, "</Types>" )
      cat(out, file = private$filename)
      self

    }
  ),
  private = list(
    filename = NULL,
    default = NULL,
    override = NULL
  )
)
