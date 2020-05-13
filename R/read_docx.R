#' @export
#' @title Create a 'Word' document object
#' @description read and import a docx file as an R object
#' representing the document. When no file is specified, it
#' uses a default empty file.
#'
#' Use then this object to add content to it and create Word files
#' from R.
#' @param x an rdocx object
#' @param path path to the docx file to use as base document.
#' @section styles:
#'
#' `read_docx()` uses a Word file as the initial document.
#' This is the original Word document from which the document
#' layout, paragraph styles, or table styles come.
#'
#' You will be able to add formatted text, change the paragraph
#' style with the R api, but it will always be easier to use
#' the styles from the original document.
#'
#' See `body_add_*` functions to add content.
#' @examples
#' # create an rdocx object with default template ---
#' read_docx()
#'
#' @seealso [print.rdocx], [body_add_par], [body_add]
read_docx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.docx")

  if(!grepl("\\.docx$", path, ignore.case = TRUE)){
    stop("read_docx only support docx files", call. = FALSE)
  }

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- structure(list( package_dir = package_dir ),
                   .Names = c("package_dir"),
                   class = "rdocx")

  obj$doc_properties <- read_core_properties(package_dir)
  obj$content_type <- content_type$new( package_dir )
  obj$doc_obj <- docx_part$new(package_dir,
                               main_file = "document.xml",
                               cursor = "/w:document/w:body/*[1]",
                               body_xpath = "/w:document/w:body")
  obj$styles <- read_docx_styles(package_dir)

  header_files <- list.files(file.path(package_dir, "word"),
                             pattern = "^header[0-9]*.xml$")
  headers <- lapply(header_files, function(x){
    docx_part$new(path = package_dir, main_file = x, cursor = "/w:hdr/*[1]", body_xpath = "/w:hdr")
  })
  names(headers) <- header_files
  obj$headers <- headers

  footer_files <- list.files(file.path(package_dir, "word"),
                             pattern = "^footer[0-9]*.xml$")
  footers <- lapply(footer_files, function(x){
    docx_part$new(path = package_dir, main_file = x, cursor = "/w:ftr/*[1]", body_xpath = "/w:ftr")
  })
  names(footers) <- footer_files
  obj$footers <- footers

  if( !file.exists(file.path(package_dir, "word", "footnotes.xml")) ){
    file.copy(system.file(package = "officer", "template", "footnotes.xml"),
              file.path(package_dir, "word", "footnotes.xml")
              )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "/word/footnotes.xml" )
    )
  }

  obj$footnotes <- docx_part$new(
    package_dir, main_file = "footnotes.xml",
    cursor = "/w:footnotes/*[last()]", body_xpath = "/w:footnotes"
  )

  default_refs <- obj$styles[obj$styles$is_default,]
  obj$default_styles <- setNames( as.list(default_refs$style_name), default_refs$style_type )

  last_sect <- xml_find_first(obj$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  obj$sect_dim <- section_dimensions(last_sect)

  obj <- cursor_end(obj)
  obj
}

#' @export
#' @param target path to the docx file to write
#' @param ... unused
#' @rdname read_docx
#' @examples
#' print(read_docx())
#' # write a rdocx object in a docx file ----
#' if( require(magrittr) ){
#'   read_docx() %>% print(target = tempfile(fileext = ".docx"))
#' }
#'
print.rdocx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("rdocx document with", length(x), "element(s)\n")
    cat("\n* styles:\n")

    style_names <- styles_info(x)
    style_sample <- style_names$style_type
    names(style_sample) <- style_names$style_name
    print(style_sample)


    cursor_elt <- x$doc_obj$get_at_cursor()
    cat("\n* Content at cursor location:\n")
    print(node_content(cursor_elt, x))
    return(invisible())
  }

  if( !grepl(x = target, pattern = "\\.(docx)$", ignore.case = TRUE) )
    stop(target , " should have '.docx' extension.")

  int_id <- 1 # unique id identifier

  # make all id unique for document
  int_id <- correct_id(x$doc_obj$get(), int_id)
  # make all id unique for footnote
  int_id <- correct_id(x$footnotes$get(), int_id)
  # make all id unique for headers
  for(docpart in x[["headers"]]){
    int_id <- correct_id(docpart$get(), int_id)
  }
  # make all id unique for footers
  for(docpart in x[["footers"]]){
    int_id <- correct_id(docpart$get(), int_id)
  }

  body <- xml_find_first(x$doc_obj$get(), "w:body")

  # If body is not ending with an sectPr, create a continuous one append it
  if( !xml_name(xml_child(body, search = xml_length(body))) %in% "sectPr" ){
    str <- paste0( "<w:sectPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                   "<w:officersection/><w:type w:val=\"continuous\"/></w:sectPr>")
    xml_add_child( body, as_xml_document(str) )
  }

  for(header in x$headers){
    header$save()
  }

  for(footer in x$footers){
    footer$save()
  }

  x <- process_sections(x)
  x$doc_obj$save()
  x$content_type$save()
  x$footnotes$save()

  # save doc properties
  x$doc_properties['modified','value'] <- format( Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
  x$doc_properties['lastModifiedBy','value'] <- Sys.getenv("USER")
  write_core_properties(x$doc_properties, x$package_dir)
  invisible(pack_folder(folder = x$package_dir, target = target ))
}

#' @export
#' @title number of blocks inside an rdocx object
#' @description return the number of blocks inside an rdocx object.
#' This number also include the default section definition of a
#' Word document - default Word section is an uninvisible element.
#' @param x an rdocx object
#' @examples
#' # how many elements are there in an new document produced
#' # with the default template.
#' length( read_docx() )
#' @family functions for Word document informations
length.rdocx <- function( x ){
  length(xml_child(x$doc_obj$get(), "w:body"))
}

#' @export
#' @title read Word styles
#' @description read Word styles and get results in
#' a tidy data.frame.
#' @param x an rdocx object
#' @param type,is_default subsets for types (i.e. paragraph) and
#' default style (when `is_default` is TRUE or FALSE)
#' @examples
#' x <- read_docx()
#' styles_info(x)
#' styles_info(x, type = "paragraph", is_default = TRUE)
#' @family functions for Word document informations
styles_info <- function( x, type = c("paragraph", "character", "table", "numbering"),
                         is_default = c(TRUE, FALSE) ){
  styles <- x$styles
  styles <- styles[styles$style_type %in% type & styles$is_default %in% is_default,]
  styles
}

#' @export
#' @title read document properties
#' @description read Word or PowerPoint document properties
#' and get results in a data.frame.
#' @param x an \code{rdocx} or \code{rpptx} object
#' @examples
#' x <- read_docx()
#' doc_properties(x)
#' @family functions for Word document informations
doc_properties <- function( x ){
  if( inherits(x, "rdocx"))
    cp <- x$doc_properties
  else if( inherits(x, "rpptx") || inherits(x, "rxlsx") ) cp <- x$core_properties
  else stop("x should be a rpptx or a rdocx or a rxlsx object.")

  out <- data.frame(tag = cp[, 'name'], value = cp[, 'value'], stringsAsFactors = FALSE)
  row.names(out) <- NULL
  out
}

#' @export
#' @title set document properties
#' @description set Word or PowerPoint document properties. These are not visible
#' in the document but are available as metadata of the document.
#' @note
#' The "last modified" and "last modified by" fields will be automatically be updated
#' when the file is written.
#' @param x an rdocx or rpptx object
#' @param title,subject,creator,description text fields
#' @param created a date object
#' @examples
#' x <- read_docx()
#' x <- set_doc_properties(x, title = "title",
#'   subject = "document subject", creator = "Me me me",
#'   description = "this document is empty",
#'   created = Sys.time())
#' x <- doc_properties(x)
#' @family functions for Word document informations
set_doc_properties <- function( x, title = NULL, subject = NULL,
                                creator = NULL, description = NULL, created = NULL ){

  if( inherits(x, "rdocx"))
    cp <- x$doc_properties
  else if( inherits(x, "rpptx")) cp <- x$core_properties
  else stop("x should be a rpptx or rdocx object.")

  if( !is.null(title) ) cp['title','value'] <- title
  if( !is.null(subject) ) cp['subject','value'] <- subject
  if( !is.null(creator) ) cp['creator','value'] <- creator
  if( !is.null(description) ) cp['description','value'] <- description
  if( !is.null(created) ) cp['created','value'] <- format( created, "%Y-%m-%dT%H:%M:%SZ")

  if( inherits(x, "rdocx"))
    x$doc_properties <- cp
  else x$core_properties <- cp

  x
}


#' @export
#' @title Word page layout
#' @description get page width, page height and margins (in inches). The return values
#' are those corresponding to the section where the cursor is.
#' @param x an \code{rdocx} object
#' @examples
#' docx_dim(read_docx())
#' @family functions for Word document informations
docx_dim <- function(x){
  cursor_elt <- x$doc_obj$get_at_cursor()
  xpath_ <- paste0(
    file.path( xml_path(cursor_elt), "following-sibling::w:sectPr"),
    "|",
    file.path( xml_path(cursor_elt), "following-sibling::w:p/w:pPr/w:sectPr"),
    "|",
    "//w:sectPr"
  )
  next_section <- xml_find_first(x$doc_obj$get(), xpath_)
  sd <- section_dimensions(next_section)
  sd$page <- sd$page / (20*72)
  sd$margins <- sd$margins / (20*72)
  sd

}


#' @export
#' @title List Word bookmarks
#' @description List bookmarks id that can be found in an \code{rdocx}
#' object.
#' @param x an \code{rdocx} object
#' @examples
#' library(magrittr)
#'
#' doc <- read_docx() %>%
#'   body_add_par("centered text", style = "centered") %>%
#'   body_bookmark("text_to_replace") %>% body_add_par("centered text", style = "centered") %>%
#'   body_bookmark("text_to_replace2")
#'
#' docx_bookmarks(doc)
#'
#' docx_bookmarks(read_docx())
#' @family functions for Word document informations
docx_bookmarks <- function(x){
  stopifnot(inherits(x, "rdocx"))

  doc_ <- xml_find_all(x$doc_obj$get(), "//w:bookmarkStart[@w:name]")
  setdiff(xml_attr(doc_, "name"), "_GoBack")
}

#' @export
#' @title replace paragraphs styles
#' @description Replace styles with others in a Word document. This function
#' is to be used for paragraph styles.
#' @param x an rdocx object
#' @param mapstyles a named list, names are the replacement style,
#' content (as a character vector) are the styles to be replaced.
#' @examples
#' library(magrittr)
#'
#' mapstyles <- list( "centered" = c("Normal"),
#'     "heading 3" = c("heading 1", "heading 2") )
#' doc <- read_docx() %>%
#'   body_add_par("A title", style = "heading 1") %>%
#'   body_add_par("Another title", style = "heading 2") %>%
#'   body_add_par("Hello world!", style = "Normal") %>%
#'   change_styles( mapstyles = mapstyles )
#'
#' print(doc, target = tempfile(fileext = ".docx"))
change_styles <- function( x, mapstyles ){

  if( is.null(mapstyles) || length(mapstyles) < 1 ) return(x)

  styles_table <- styles_info(x, type = "paragraph")

  from_styles <- unique( as.character( unlist(mapstyles) ) )
  to_styles <- unique( names( mapstyles) )

  if( any( is.na( mfrom <- match( from_styles, styles_table$style_name ) ) ) ){
    stop("could not find paragraph style ", paste0( shQuote(from_styles[is.na(mfrom)]), collapse = ", " ), ".", call. = FALSE)
  }
  if( any( is.na( mto <- match( to_styles, styles_table$style_name ) ) ) ){
    stop("could not find paragraph style ", paste0( shQuote(to_styles[is.na(mto)]), collapse = ", " ), ".", call. = FALSE)
  }

  mapping <- mapply(function(from, to) {
    id_to <- which( styles_table$style_name %in% to )
    id_to <- styles_table$style_id[id_to]

    id_from <- which( styles_table$style_name %in% from )
    id_from <- styles_table$style_id[id_from]

    data.frame( from = id_from, to = rep(id_to, length(from)), stringsAsFactors = FALSE )
  }, mapstyles, names(mapstyles), SIMPLIFY = FALSE)

  mapping <- do.call(rbind, mapping)
  row.names(mapping) <- NULL

  for(i in seq_len( nrow(mapping) )){
    all_nodes <- xml_find_all(x$doc_obj$get(), sprintf("//w:pStyle[@w:val='%s']", mapping$from[i]))
    xml_attr(all_nodes, "w:val") <- rep(mapping$to[i], length(all_nodes) )
  }

  x
}



#' @export
#' @title body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_xml(doc)
docx_body_xml <- function( x ){
  x$doc_obj$get()
}

#' @export
#' @title body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_relationship(doc)
docx_body_relationship <- function( x ){
  x$doc_obj$relationship()
}

