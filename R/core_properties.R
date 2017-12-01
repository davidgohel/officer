core_properties <- R6Class(
  "core_properties",
  public = list(

    initialize = function(package_dir) {
      private$filename <- file.path(package_dir, "docProps/core.xml")
      private$doc <- read_xml(private$filename)

    },
    get_data = function() {
      all_ <- xml_find_all(private$doc, "/cp:coreProperties/*")
      data.frame(stringsAsFactors = FALSE,
        tag = xml_name(all_),
        value = xml_text(all_)
      )
    },
    set_title = function(value){
      private$set_core_property( "title", "dc", value )
      self
    },
    set_subject = function(value){
      private$set_core_property( "subject", "dc", value )
      self
    },
    set_creator = function(value){
      private$set_core_property( "creator", "dc", value )
      self
    },
    set_keywords = function(value){
      private$set_core_property( "keywords", "cp", value )
      self
    },
    set_description = function(value){
      private$set_core_property( "description", "dc", value )
      self
    },
    set_modified_by = function(value){
      private$set_core_property( "lastModifiedBy", "cp", value )
      self
    },
    set_last_modified = function(value){
      private$set_core_property( "modified", "dcterms", value, c("xsi:type"="dcterms:W3CDTF") )
      self
    },
    set_created = function(value){
      private$set_core_property( "created", "dcterms", value, c("xsi:type"="dcterms:W3CDTF") )
      self
    },
    save = function() {
      write_xml(private$doc, file = private$filename)
      self
    }
  ),
  private = list(
    filename = NULL,
    doc = NULL,

    set_core_property = function( tag, ns, value, attrs = NULL ) {
      ns_list <- c(cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                   dc="http://purl.org/dc/elements/1.1/",
                   dcterms="http://purl.org/dc/terms/",
                   dcmitype="http://purl.org/dc/dcmitype/"
                   )
      stopifnot(ns %in% names(ns_list) )
      if(is.null(attrs))
        str <- sprintf("<%s:%s xmlns:%s=\"%s\">%s</%s:%s>", ns, tag, ns, ns_list[ns], value, ns, tag)
      else
        str <- sprintf("<%s:%s %s xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:%s=\"%s\">%s</%s:%s>", ns, tag,
                       paste0(names(attrs), "=", shQuote(attrs, type = "cmd"), collapse = " "),
                       ns, ns_list[ns], value, ns, tag)
      obj <- as_xml_document(str)
      node <- xml_find_first(private$doc, sprintf("/cp:coreProperties/%s:%s", ns, tag))
      if( !inherits(node, "xml_missing"))
        xml_replace(node, obj)
      else
        xml_add_child(xml_find_first(private$doc, "/cp:coreProperties"), obj)
      self
    }
  )
)

