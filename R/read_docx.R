#' @export
#' @title open a connexion to a 'Word' file
#' @description read and import a docx file as an R object
#' representing the document.
#' @param path path to the docx file to use a base document.
#' @param x a rdocx object
#' @examples
#' # create a rdocx object with default template ---
#' read_docx()
#'
#' @importFrom xml2 read_xml xml_length xml_find_first as_list
read_docx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.docx")

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- structure(list( package_dir = package_dir ),
                   .Names = c("package_dir"),
                   class = "rdocx")

  obj$content_type <- content_type$new( obj )
  obj$doc_obj <- docx_document$new(package_dir)

  default_refs <- styles_info(obj)
  default_refs <- default_refs[default_refs$is_default,]
  obj$default_styles <- setNames( as.list(default_refs$style_name), default_refs$style_type )

  last_sect <- xml_find_first(obj$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  section_obj <- as_list(last_sect)
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
#'   read_docx() %>% print(target = "out.docx")
#'   # full path of produced file is returned
#'   print(.Last.value)
#' }
#'
#' @importFrom xml2 xml_attr<- xml_find_all xml_find_all
#' @importFrom purrr walk2
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

  # make all id unique
  all_uid <- xml_find_all(x$doc_obj$get(), "//*[@id]")
  walk2(all_uid, seq_along(all_uid), function(x, z) {
    xml_attr(x, "id") <- z
    x
  })

  sections_ <- xml_find_all(x$doc_obj$get(), "//w:sectPr")
  last_sect <- sections_[length(sections_)]
  if( inherits( xml_find_first(x$doc_obj$get(), file.path( xml_path(last_sect), "w:type")), "xml_missing" ) ){
    xml_add_child( last_sect,
      as_xml_document("<w:type xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:val=\"continuous\"/>")
      )
  }
  x$doc_obj$save()
  x$content_type$save()

  # save doc properties
  x$doc_obj$get_doc_properties()$set_last_modified(format( Sys.time(), "%Y-%m-%dT%H:%M:%SZ"))
  x$doc_obj$get_doc_properties()$set_modified_by(Sys.getenv("USER"))
  x$doc_obj$get_doc_properties()$save()

  pack_folder(folder = x$doc_obj$package_dirname(), target = target )
}

#' @export
#' @examples
#' # how many element are there in the document ----
#' length( read_docx() )
#'
#' @importFrom xml2 read_xml xml_length xml_find_first
#' @rdname read_docx
length.rdocx <- function( x ){
  xml_find_first(x$doc_obj$get(), "/w:document/w:body") %>% xml_length()

}

#' @export
#' @title read Word styles
#' @description read Word styles and get results in
#' a tidy data.frame.
#' @param x a rdocx object
#' @examples
#' library(magrittr)
#' read_docx() %>% styles_info()
styles_info <- function( x ){
  x$doc_obj$styles()
}

#' @export
#' @title read document properties
#' @description read Word or PowerPoint document properties
#' and get results in a tidy data.frame.
#' @param x an \code{rdocx} or \code{rpptx} object
#' @examples
#' library(magrittr)
#' read_docx() %>% doc_properties()
doc_properties <- function( x ){
  if( inherits(x, "rdocx"))
    cp <- x$doc_obj$get_doc_properties()
  else if( inherits(x, "rpptx")) cp <- x$core_properties
  else stop("x should be a rpptx or rdocx object.")

  cp$get_data()
}

#' @export
#' @title set document properties
#' @description set Word or PowerPoint document properties. These are not visible
#' in the document but are available as metadata of the document.
#' @note
#' Fields "last modified" and "last modified by" will be automatically be updated
#' when file will be written.
#' @param x a rdocx or rpptx object
#' @param title,subject,creator,description text fields
#' @param created a date object
#' @examples
#' library(magrittr)
#' read_docx() %>% set_doc_properties(title = "title",
#'   subject = "document subject", creator = "Me me me",
#'   description = "this document is empty",
#'   created = Sys.time()) %>% doc_properties()
set_doc_properties <- function( x, title = NULL, subject = NULL,
                                creator = NULL, description = NULL, created = NULL ){

  if( inherits(x, "rdocx"))
    cp <- x$doc_obj$get_doc_properties()
  else if( inherits(x, "rpptx")) cp <- x$core_properties
  else stop("x should be a rpptx or rdocx object.")

  if( !is.null(title) ) cp$set_title(title)
  if( !is.null(subject) ) cp$set_subject(subject)
  if( !is.null(creator) ) cp$set_creator(creator)
  if( !is.null(description) ) cp$set_description(description)
  if( !is.null(created) ) cp$set_created(format( created, "%Y-%m-%dT%H:%M:%SZ"))

  x
}


#' @export
#' @title Word page layout
#' @description get page width, page height and margins (in inches). The return values
#' are those corresponding to the section where the cursor is.
#' @param x a \code{rdocx} object
#' @examples
#' docx_dim(read_docx())
docx_dim <- function(x){
  cursor_elt <- x$doc_obj$get_at_cursor()
  xpath_ <- file.path( xml_path(cursor_elt), "following-sibling::w:sectPr")
  next_section <- xml_find_first(x$doc_obj$get(), xpath_)
  sd <- section_dimensions(next_section)
  sd$page <- sd$page / 914400
  sd$margins <- sd$margins / 914400
  sd

}
