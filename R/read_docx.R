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
#' @importFrom xml2 read_xml xml_length xml_find_first
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

  default_refs <- filter_(styles_info(obj) , interp(~ is_default) )
  obj$default_styles <- setNames( as.list(default_refs$style_name), default_refs$style_type )

  obj <- cursor_end(obj)
  obj
}

#' @export
#' @param target path to the docx file to write
#' @param ... unused
#' @rdname read_docx
#' @examples
#' # write a rdocx object in a docx file ----
#' if( require(magrittr) && has_zip() ){
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
    cat("Available styles are:\n")
    print(as.data.frame(select_( styles_info(x), "style_type", "style_name")))
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
  x$doc_obj$save()
  x$content_type$save()

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


