#' @importFrom xml2 xml_find_all xml_attr read_xml
#' @import magrittr
#' @importFrom xml2 xml_ns read_xml xml_find_all xml_name xml_text xml_text<- xml_remove
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
      private$doc_properties <- core_properties$new(private$package_dir)
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
      xml_length( xml_find_first(self$get(), "/w:document/w:body") )

    },

    get_at_cursor = function() {
      node <- xml_find_first(self$get(), private$cursor)
      if( inherits(node, "xml_missing") )
        stop("cursor does not correspond to any node", call. = FALSE)
      node
    },

    get_doc_properties = function(){
      private$doc_properties
    },
    set_cursor = function( cursor ){
      private$cursor <- cursor
      self
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

    cursor_bookmark = function( id ){
      xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", id)
      bm_start <- xml_find_first(self$get(), xpath_)

      if( inherits(bm_start, "xml_missing") )
        stop("cannot find bookmark ", shQuote(id), call. = FALSE)

      bm_id <- xml_attr(bm_start, "id")

      xpath_ <- sprintf("/w:document/w:body/*[w:bookmarkStart[@w:id='%s']]", bm_id)
      par_with_bm <- xml_find_first(self$get(), xpath_)

      cursor <- xml_path(par_with_bm)

      xpath_ <- paste0( cursor, sprintf("/w:bookmarkEnd[@w:id='%s']", bm_id) )
      nodes_with_bm_end <- xml_find_first(self$get(), xpath_)
      if( inherits(nodes_with_bm_end, "xml_missing") )
        stop("bookmark ", shQuote(id), " does not end in the same paragraph (or is on the whole paragraph)", call. = FALSE)

      cursor <- cursor
      private$cursor <- cursor
      self
    },

    cursor_replace_first_text = function( id, text ){

      xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", id)
      bm_start <- xml_find_first(self$get(), xpath_)
      if( inherits(bm_start, "xml_missing") )
        stop("cannot find bookmark ", shQuote(id), call. = FALSE)

      str_ <- sprintf("//w:bookmarkStart[@w:name='%s']/following-sibling::w:r", id )
      following_start <- sapply( xml_find_all(self$get(), str_), xml_path )
      str_ <- sprintf("//w:bookmarkEnd[@w:id='%s']/preceding-sibling::w:r", xml_attr(bm_start, "id") )
      preceding_end <- sapply( xml_find_all(self$get(), str_), xml_path )

      match_path <- base::intersect(following_start, preceding_end)
      if( length(match_path) < 1 )
        stop("could not find any bookmark ", id, " located INSIDE a single paragraph" )

      run_nodes <- xml_find_all(self$get(), paste0( match_path, collapse = "|" ) )

      for(node in run_nodes[setdiff(seq_along(run_nodes), 1)])
        xml_remove(node)

      xml_text(run_nodes[[1]] ) <- text
      self
    },

    replace_all_text = function( oldValue, newValue, onlyAtCursor=TRUE, ... ) {

      replacement_count <- 0

      base_node <- if (onlyAtCursor) self$get_at_cursor() else self$get()

      # For each matching text node...
      for (text_node in xml_find_all(base_node, ".//w:t")) {
        # ...if it contains the oldValue...
        if (grepl(oldValue, xml_text(text_node), ...)) {
          replacement_count <- replacement_count + 1
          # Replace the node text with the newValue.
          xml_text(text_node) <- gsub(oldValue, newValue, xml_text(text_node), ...)
        }
      }

      # Alert the user if no replacements were made.
      if (replacement_count == 0) {
        search_zone_text <- if (onlyAtCursor) "at the cursor." else "in the document."
        warning("Found 0 instances of '", oldValue, "' ", search_zone_text)
      }

      self
    },

    docx_show_chunk = function() {
      # Show the structure of how the text is split along `<w:t>` tags at the
      # current cursor.
      text_nodes <- xml_find_all(self$get_at_cursor(), ".//w:t")
      message(length(text_nodes), " text nodes found at this cursor.")
      for (text_node in text_nodes) {
        message("  <w:t>: '", xml_text(text_node), "'")
      }

      self
    },

    cursor_reach = function( keyword ){
      nodes_with_text <- xml_find_all(self$get(),"/w:document/w:body/*[.//*/text()]")

      if( length(nodes_with_text) < 1 )
        stop("no text found in the document", call. = FALSE)

      text_ <- xml_text(nodes_with_text)
      test_ <- grepl(pattern = keyword, x = text_)
      if( !any(test_) )
        stop(keyword, " has not been found in the document", call. = FALSE)

      node <- nodes_with_text[[ which(test_)[1] ]]
      private$cursor <- xml_path(node)
      self
    },

    cursor_forward = function( ){
      xpath_ <- paste0(private$cursor, "/following-sibling::*" )
      private$cursor <- xml_path( xml_find_first(self$get(), xpath_ ) )
      self
    },

    cursor_backward = function( ){
      xpath_ <- paste0(private$cursor, "/preceding-sibling::*[1]" )
      private$cursor <- xml_path( xml_find_first(self$get(), xpath_ ) )
      self
    }

  ),
  private = list(
    package_dir = NULL,
    cursor = NULL,
    styles_df = NULL,
    doc_properties = NULL,


    read_styles = function(  ){
      styles_file <- file.path(private$package_dir, "word/styles.xml")
      doc <- read_xml(styles_file)

      all_styles <- xml_find_all(doc, "/w:styles/w:style")
      all_desc <- data.frame(stringsAsFactors = FALSE,
        style_type = xml_attr(all_styles, "type"),
        style_id = xml_attr(all_styles, "styleId"),
        style_name = xml_attr(xml_find_all(all_styles, "w:name"), "val"),
        is_custom = xml_attr(all_styles, "customStyle") %in% "1",
        is_default = xml_attr(all_styles, "default") %in% "1"
      )

      all_desc
    }

  )

)
