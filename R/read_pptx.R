#' @export
#' @title Create a 'PowerPoint' document object
#' @description Read and import a pptx file as an R object
#' representing the document.
#'
#' The function is called `read_pptx` because it allows you to initialize an
#' object of class `rpptx` from an existing PowerPoint file. Content will be
#' added to the existing presentation. By default, an empty document is used.
#'
#' @param path path to the pptx file to use as base document. `potx` file are supported.
#' @section master layouts and slide layouts:
#' `read_pptx()` uses a PowerPoint file as the initial document. This is the original
#' PowerPoint document where all slide layouts, placeholders for shapes and
#' styles come from. Major points to be aware of are:
#'
#'   * Slide layouts are relative to a master layout. A document can contain one or
#' more master layouts; a master layout can contain one or more slide layouts.
#' * A slide layout inherits design properties from its master layout but some
#' properties can be overwritten.
#' * Designs and formatting properties of layouts and shapes (placeholders in a
#'                                                            layout) are defined within the initial document. There is no R function to
#' modify these values - they must be defined in the initial document.
#' @examples
#' read_pptx()
#' @seealso [print.rpptx()], [add_slide()], [plot_layout_properties()], [ph_with()]
read_pptx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.pptx")

  if(!grepl("\\.(pptx|potx)$", path, ignore.case = TRUE)){
    stop("read_pptx only support pptx files", call. = FALSE)
  }

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- list(package_dir = package_dir)


  obj$table_styles <- read_table_style(package_dir)

  obj$presentation <- presentation$new(package_dir)

  obj$masterLayouts <- dir_master$new(package_dir, slide_master$new("ppt/slideMasters") )

  obj$slideLayouts <- dir_layout$new( package_dir,
                                      master_metadata = obj$masterLayouts$get_metadata(),
                                      master_xfrm = obj$masterLayouts$xfrm() )

  obj$slide <- dir_slide$new( package_dir, obj$slideLayouts$get_xfrm_data() )
  obj$notesSlide <- dir_notesSlide$new(package_dir)
  obj$notesMaster <- dir_notesMaster$new(package_dir, slide_master$new("ppt/notesMasters"))
  obj$content_type <- content_type$new( package_dir )
  obj$core_properties <- read_core_properties(package_dir)
  obj$doc_properties_custom <- read_custom_properties(package_dir)

  obj$rel <- relationship$new()
  obj$rel$feed_from_xml(file.path(package_dir, "_rels", ".rels"))

  obj$cursor = obj$slide$length()
  obj$layout_default <- list(layout = NA, master = NA)

  class(obj) <- "rpptx"
  obj
}


read_table_style <- function(path){
  file <- file.path(path, "ppt/tableStyles.xml")
  if (!file.exists(file)) {
    warning("tableStyles.xml file does not exist in PPTX")
    return(NULL)
  }
  doc <- read_xml(file)
  nodes <- xml_find_all(doc, "//a:tblStyleLst")
  data.frame(def = xml_attr(nodes, "def"),
             styleName = xml_attr(nodes, "styleName"),
             stringsAsFactors = FALSE )
}


#' @title Write a 'PowerPoint' file.
#' @description Create a 'PowerPoint' file from an `rpptx` object (created by [read_pptx()]).
#' @param x an `rpptx` object.
#' @param target path to the .pptx file to write. If `target` is `NULL` (default), the `rpptx` object is printed to
#'   the console.
#' @param preview Save `x` to a temporary file and open it (default `FALSE`).
#' @param ... unused.
#' @returns If preview is `TRUE`, returns the temp file path invisibly.
#' @examples
#' # write an rpptx object to a .pptx file ----
#' file <- tempfile(fileext = ".pptx")
#' x <- read_pptx() # empty presentation, has no slides yet
#' print(x, target = file)
#'
#' # preview mode: save to temp file and open locally ----
#' \dontrun{
#' print(x, preview = TRUE)
#' }
#' @export
#' @seealso [read_pptx()]
print.rpptx <- function(x, target = NULL, preview = FALSE, ...) {
  if (preview) {
    file <- tempfile(fileext = ".pptx")
    print.rpptx(x, target = file, preview = FALSE)
    open_file(file)
    return(invisible(file))
  }
  if (is.null(target)) {
    cli::cli_text("pptx document with {.val {length(x)}} slide{?s}")
    cli::cli_text("Available layouts and their associated master(s):")
    df <- as.data.frame(layout_summary(x))
    if (has_layout_default(x)) {
      cli::cli_text(cli::col_grey("(*) = Default layout"))
      la <- get_layout_default(x)
      i <- which(df$layout == la$layout & df$master == la$master)
      df$" " <- " " # empty column name
      df$" "[i] <- "*"
    }
    print(df)
    return(invisible())
  }

  if (!grepl(x = target, pattern = "\\.(pptx)$", ignore.case = TRUE)) {
    cli::cli_abort("{.val {target}} should have '.pptx' extension.")
  }

  if (is_windows() && is_doc_open(target)) {
    cli::cli_abort("{.val {target}} is open. To write to this document, please, close it.")
  }
  x <- pptx_fortify_slides(x)
  x$rel$write(file.path(x$package_dir, "_rels", ".rels"))

  # viewProps - drop lastView if available
  viewProps <- read_xml(file.path(x$package_dir, "ppt/viewProps.xml"))
  if (!is.na(xml_attr(viewProps, "lastView"))) {
    xml_attr(viewProps, "lastView") <- NULL
  }
  write_xml(viewProps, file.path(x$package_dir, "ppt/viewProps.xml"))

  drop_templatenode_from_app(x$package_dir)

  x$presentation$save()
  x$content_type$save()
  x$slide$save_slides()
  x$notesSlide$save_slides()

  x$core_properties["modified", "value"] <- format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
  x$core_properties["lastModifiedBy", "value"] <- Sys.getenv("USER")
  write_core_properties(x$core_properties, x$package_dir)

  if (nrow(x$doc_properties_custom$data) > 0) {
    write_custom_properties(x$doc_properties_custom, x$package_dir)
  }

  x <- sanitize_images(x, warn_user = FALSE)
  invisible(pack_folder(folder = x$package_dir, target = target))
}

