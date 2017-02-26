#' @importFrom xml2 xml_find_all xml_attr read_xml
#' @import magrittr
#' @importFrom tibble tibble
#' @importFrom xml2 xml_ns read_xml xml_find_all xml_name xml_text
docx_document <- R6Class(
  "docx_document",
  inherit = openxml_document,
  public = list(

    initialize = function( path ) {
      super$initialize("word")
      private$package_dir <- path
      super$feed(file.path(private$package_dir, "word/document.xml"))
      private$cursor <- "/w:document/w:body/*[1]"
      private$styles_df <- private$read_styles()
    },

    package_dirname = function(){
      private$package_dir
    },

    styles = function(){
      private$styles_df
    },
    get_style_id = function(style, type ){
      ref <- private$styles_df[private$styles_df$style_type==type, ]
      if(!style %in% ref$style_name){
        t_ <- shQuote(ref$style_name, type = "sh")
        t_ <- paste(t_, collapse = ", ")
        t_ <- paste0("c(", t_, ")")
        stop("could not match any style named ", shQuote(style, type = "sh"), " in ", t_, call. = FALSE)
      }
      ref$style_id[ref$style_name == style]
    },

    length = function( ){
      xml_find_first(self$get(), "/w:document/w:body") %>% xml_length()

    },

    get_at_cursor = function() {
      node <- xml_find_first(self$get(), private$cursor)
      if( inherits(node, "xml_missing") )
        stop("cursor does not correspond to any node", call. = FALSE)
      node
    },


    cursor_begin = function( ){
      private$cursor <- "/w:document/w:body/*[1]"
      self
    },

    cursor_end = function( ){
      len <- self$length()
      if( len < 2 ) private$cursor <- "/w:document/w:body/*[1]"
      else private$cursor <- sprintf("/w:document/w:body/*[%.0f]", len - 1 )
      self
    },

    cursor_reach = function( keyword ){
      xpath_ <- sprintf("/w:document/w:body/*[contains(.//*/w:t/text(),'%s')]", keyword)
      cursor <- xml_find_first(self$get(), xpath_) %>% xml_path()
      if( inherits(cursor, "xml_missing") )
        stop(keyword, " has not been found in the document", call. = FALSE)
      private$cursor <- cursor
      self
    },

    cursor_forward = function( ){
      xpath_ <- paste0(private$cursor, "/following-sibling::*" )
      private$cursor <- xml_find_first(self$get(), xpath_ ) %>% xml_path()
      self
    },


    cursor_backward = function( ){
      xpath_ <- paste0(private$cursor, "/preceding-sibling::*[1]" )
      private$cursor <- xml_find_first(self$get(), xpath_ ) %>% xml_path()
      self
    }

  ),
  private = list(
    package_dir = NULL,
    cursor = NULL,
    styles_df = NULL,


    read_styles = function(  ){
      styles_file <- file.path(private$package_dir, "word/styles.xml")
      doc <- read_xml(styles_file)

      all_styles <- xml_find_all(doc, "/w:styles/w:style")
      all_desc <- tibble(
        style_type = all_styles %>% xml_attr("type"),
        style_id = all_styles %>% xml_attr("styleId"),
        style_name = all_styles %>% xml_find_all("w:name") %>% xml_attr("val"),
        is_custom = all_styles %>% xml_attr("customStyle") %in% "1",
        is_default = all_styles %>% xml_attr("default") %in% "1"
      )

      all_desc
    }
    # , read_core = function(  ){
    #   core_file <- file.path(private$package_dir, "docProps/core.xml")
    #   doc <- read_xml(core_file)
    #
    #   all_ <- xml_find_all(doc, "/cp:coreProperties/*")
    #   ns_ <- xml_ns(doc)
    #   all_desc <- tibble(
    #     tag_ns = all_ %>% xml_name(ns = ns_),
    #     tag = all_ %>% xml_name(),
    #     value = all_ %>% xml_text()
    #   )
    #
    #   all_desc
    # }

  )

)
