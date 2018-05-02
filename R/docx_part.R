#' @importFrom xml2 xml_find_all xml_attr read_xml
#' @import magrittr
#' @importFrom xml2 xml_ns read_xml xml_find_all xml_name xml_text xml_text<- xml_remove
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
    },

    get_at_cursor = function() {
      node <- xml_find_first(self$get(), private$cursor)
      if( inherits(node, "xml_missing") )
        stop("cursor does not correspond to any node", call. = FALSE)
      node
    },

    set_cursor = function( cursor ){
      private$cursor <- cursor
      self
    },
    cursor_begin = function( ){
      private$cursor <- paste0(private$body_xpath, "/*[1]")
      self
    },

    cursor_end = function( ){
      len <- self$length()
      if( len < 2 ) private$cursor <- paste0(private$body_xpath, "/*[1]")
      else private$cursor <- sprintf(paste0(private$body_xpath, "/*[%.0f]"), len - 1 )
      self
    },

    cursor_bookmark = function( id ){
      xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", id)
      bm_start <- xml_find_first(self$get(), xpath_)

      if( inherits(bm_start, "xml_missing") )
        stop("cannot find bookmark ", shQuote(id), call. = FALSE)

      bm_id <- xml_attr(bm_start, "id")


      xpath_ <- sprintf(paste0(private$body_xpath, "//*[w:bookmarkStart[@w:id='%s']]"), bm_id)
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

    has_bookmark = function( id ){
      xpath_ <- sprintf("//w:bookmarkStart[@w:name='%s']", id)
      bm_start <- xml_find_first(self$get(), xpath_)

      if( inherits(bm_start, "xml_missing") )
        FALSE
      else TRUE
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

    cursor_replace_first_img = function( id, src, width, height ){

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

      new_src <- tempfile( fileext = gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", src) )
      file.copy( src, to = new_src )

      blip_id <- self$relationship()$get_next_id()
      self$relationship()$add_img(new_src, root_target = "media")

      img_path <- file.path(private$package_dir, "word", "media")
      dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
      file.copy(from = new_src, to = file.path(private$package_dir, "word", "media", basename(new_src)))

      out <- wml_image(paste0("rId", blip_id), width = width*72, height = height*72)

      xml_replace(run_nodes[[1]], as_xml_document(out) )
      self
    },

    replace_all_text = function( oldValue, newValue, onlyAtCursor=TRUE, warn = TRUE, ... ) {

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
      if (replacement_count == 0 && warn) {
        search_zone_text <- if (onlyAtCursor) "at the cursor." else "in the document."
        warning("Found 0 instances of '", oldValue, "' ", search_zone_text)
      }

      self
    },

    docx_show_chunk = function() {
      # Show the structure of how the text is split along `<w:t>` tags at the
      # current cursor.
      text_nodes <- xml_find_all(self$get_at_cursor(), ".//w:t")
      msg <- paste0(length(text_nodes), " text nodes found at this cursor.")
      msg_detail <- ""
      for (text_node in text_nodes) {
        msg_detail <- paste0( msg_detail,
                         paste0("\n  <w:t>: '",
                                xml_text(text_node), "'") )
      }
      message(paste(msg, msg_detail))
      self
    },

    cursor_reach = function( keyword ){

      nodes_with_text <- xml_find_all(
        self$get(),
        paste0(private$body_xpath, "/*[.//*/text()]")
        )

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
      node_at_cursor <- xml_find_first(self$get(), xpath_ )
      if( inherits(node_at_cursor, "xml_missing") )
        stop("cannot move forward the cursor as there is no more content in this area", call. = FALSE)
      private$cursor <- xml_path( node_at_cursor )
      self
    },

    cursor_backward = function( ){
      xpath_ <- paste0(private$cursor, "/preceding-sibling::*[1]" )
      node_at_cursor <- xml_find_first(self$get(), xpath_ )
      if( inherits(node_at_cursor, "xml_missing") )
        stop("cannot move backward the cursor as there is no previous content in this area", call. = FALSE)
      private$cursor <- xml_path( node_at_cursor )
      self
    }

  ),
  private = list(
    package_dir = NULL,
    cursor = NULL,
    body_xpath = NULL
  )

)
