#' @export
#' @title Create a 'Word' document object
#' @description read and import a docx file as an R object
#' representing the document. When no file is specified, it
#' uses a default empty file.
#'
#' Use then this object to add content to it and create Word files
#' from R.
#' @param path path to the docx file to use as base document.
#' `dotx` file are supported.
#' @return an object of class `rdocx`.
#' @section styles:
#'
#' `read_docx()` uses a Word file as the initial document.
#' This is the original Word document from which the document
#' layout, paragraph styles, or table styles come.
#'
#' You will be able to add formatted text, change the paragraph
#' style with the R api but also use the styles from the
#' original document.
#'
#' See `body_add_*` functions to add content.
#' @examples
#' library(officer)
#'
#' pinst <- plot_instr({
#'   z <- c(rnorm(100), rnorm(50, mean = 5))
#'   plot(density(z))
#' })
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, "This is a table", style = "heading 2")
#' doc_1 <- body_add_table(doc_1, value = mtcars, style = "table_template")
#' doc_1 <- body_add_par(doc_1, "This is a plot", style = "heading 2")
#' doc_1 <- body_add_plot(doc_1, pinst)
#' docx_file_1 <- print(doc_1, target = tempfile(fileext = ".docx"))
#'
#' template <- system.file(package = "officer",
#'   "doc_examples", "landscape.docx")
#' doc_2 <- read_docx(path = template)
#' doc_2 <- body_add_par(doc_2, "This is a table", style = "heading 2")
#' doc_2 <- body_add_table(doc_2, value = mtcars)
#' doc_2 <- body_add_par(doc_2, "This is a plot", style = "heading 2")
#' doc_2 <- body_add_plot(doc_2, pinst)
#' docx_file_2 <- print(doc_2, target = tempfile(fileext = ".docx"))
#'
#' @seealso [body_add_par], [body_add_plot], [body_add_table]
#' @section Illustrations:
#'
#' \if{html}{\figure{read_docx_doc_1.png}{options: width=80\%}}
#'
#' \if{html}{\figure{read_docx_doc_2.png}{options: width=80\%}}
read_docx <- function(path = NULL) {
  if (!is.null(path) && !file.exists(path)) {
    stop("could not find file ", shQuote(path), call. = FALSE)
  }

  if (is.null(path)) {
    path <- system.file(package = "officer", "template/template.docx")
  }

  if (!grepl("\\.(docx|dotx)$", path, ignore.case = TRUE)) {
    stop("read_docx only support docx files", call. = FALSE)
  }

  package_dir <- tempfile()
  unpack_folder(file = path, folder = package_dir)

  obj <- structure(list(package_dir = package_dir),
    .Names = c("package_dir"),
    class = "rdocx"
  )

  obj$settings <- update(
    object = docx_settings(),
    file = file.path(package_dir, "word", "settings.xml")
  )

  obj$rel <- relationship$new()
  obj$rel$feed_from_xml(file.path(package_dir, "_rels", ".rels"))

  obj$doc_properties_custom <- read_custom_properties(package_dir)
  obj$doc_properties <- read_core_properties(package_dir)
  obj$content_type <- content_type$new(package_dir)
  obj$doc_obj <- docx_part$new(package_dir,
    main_file = "document.xml",
    cursor = "/w:document/w:body/*[1]",
    body_xpath = "/w:document/w:body"
  )
  obj$styles <- read_docx_styles(package_dir)
  obj$officer_cursor <- officer_cursor(obj$doc_obj$get())

  obj$headers <- update_hf_list(part_list = list(), type = "header", package_dir = package_dir)
  obj$footers <- update_hf_list(part_list = list(), type = "footer", package_dir = package_dir)

  if (!file.exists(file.path(package_dir, "word", "footnotes.xml"))) {
    file.copy(
      system.file(package = "officer", "template", "footnotes.xml"),
      file.path(package_dir, "word", "footnotes.xml")
    )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "/word/footnotes.xml")
    )
  }

  obj$footnotes <- docx_part$new(
    package_dir,
    main_file = "footnotes.xml",
    cursor = "/w:footnotes/*[last()]", body_xpath = "/w:footnotes"
  )

  default_refs <- obj$styles[obj$styles$is_default, ]
  obj$default_styles <- setNames(as.list(default_refs$style_name), default_refs$style_type)

  last_sect <- xml_find_first(obj$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  obj$sect_dim <- section_dimensions(last_sect)

  obj <- cursor_end(obj)
  obj
}

#' @export
#' @describeIn read_docx write docx to a file. It returns the path of the result
#' file.
#' @param x an rdocx object
#' @param target path to the docx file to write
#' @param ... unused
print.rdocx <- function(x, target = NULL, ...) {
  if (is.null(target)) {
    cat("rdocx document with", length(x), "element(s)\n")
    cat("\n* styles:\n")

    style_names <- styles_info(x)
    style_sample <- style_names$style_type
    names(style_sample) <- style_names$style_name
    print(style_sample)

    if (length(x) > 1) {
      cursor_elt <- docx_current_block_xml(x)
      cat("\n* Content at cursor location:\n")
      print(node_content(cursor_elt, x))
    } else {
      cat("\n* empty document\n")
    }
    return(invisible())
  }

  if (!grepl(x = target, pattern = "\\.(docx)$", ignore.case = TRUE)) {
    stop(target, " should have '.docx' extension.")
  }

  if (is_windows() && is_office_doc_edited(target)) {
    stop(
      target, " is already edited.",
      " You must close the document in order to be able to write the file."
    )
  }

  x <- process_sections(x)

  process_footnotes(x)
  process_stylenames(x$doc_obj, x$styles)
  process_links(x$doc_obj, type = "wml")
  process_docx_poured(
    doc_obj = x$doc_obj,
    relationships = x$doc_obj$relationship(),
    content_type = x$content_type,
    package_dir = x$package_dir
  )
  process_images(x$doc_obj, x$doc_obj$relationship(), x$package_dir)
  process_images(x$footnotes, x$footnotes$relationship(), x$package_dir)

  x$headers <- update_hf_list(part_list = x$headers, type = "header", package_dir = x$package_dir)
  x$footers <- update_hf_list(part_list = x$footers, type = "footer", package_dir = x$package_dir)
  for (header in x$headers) process_links(header, type = "wml")
  for (footer in x$footers) process_links(footer, type = "wml")
  for (header in x$headers) process_images(header, header$relationship(), x$package_dir)
  for (footer in x$footers) process_images(footer, footer$relationship(), x$package_dir)

  int_id <- 1 # unique id identifier

  # make all id unique for document
  int_id <- correct_id(x$doc_obj$get(), int_id)
  # make all id unique for footnote
  int_id <- correct_id(x$footnotes$get(), int_id)
  # make all id unique for headers
  for (docpart in x[["headers"]]) {
    int_id <- correct_id(docpart$get(), int_id)
  }
  # make all id unique for footers
  for (docpart in x[["footers"]]) {
    int_id <- correct_id(docpart$get(), int_id)
  }

  body <- xml_find_first(x$doc_obj$get(), "w:body")

  # If body is not ending with an sectPr, create a continuous one append it
  if (!xml_name(xml_child(body, search = xml_length(body))) %in% "sectPr") {
    str <- paste0(
      "<w:sectPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
      "<w:type w:val=\"continuous\"/></w:sectPr>"
    )
    xml_add_child(body, as_xml_document(str))
  }

  for (header in x$headers) {
    header$save()
  }
  for (footer in x$footers) {
    footer$save()
  }
  x$doc_obj$save()
  x$content_type$save()
  x$footnotes$save()

  x$rel$write(file.path(x$package_dir, "_rels", ".rels"))
  write_docx_settings(x)

  # save doc properties
  if (nrow(x$doc_properties$data) > 0) {
    x$doc_properties["modified", "value"] <- format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
    x$doc_properties["lastModifiedBy", "value"] <- Sys.getenv("USER")
    write_core_properties(x$doc_properties, x$package_dir)
  }
  if (nrow(x$doc_properties_custom$data) > 0) {
    write_custom_properties(x$doc_properties_custom, x$package_dir)
  }
  invisible(pack_folder(folder = x$package_dir, target = target))
}



#' @export
#' @title Number of blocks inside an rdocx object
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
  xml_length(xml_child(x$doc_obj$get(), "w:body"))
}

#' @export
#' @title Read 'Word' styles
#' @description read Word styles and get results in
#' a data.frame.
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
#' @title Add or replace paragraph style in a Word document
#' @description The function lets you add or replace a Word paragraph style.
#' @param x an rdocx object
#' @param style_id a unique style identifier for Word.
#' @param style_name a unique label associated with the style identifier.
#' This label is the name of the style when Word edit the document.
#' @param base_on the style name used as base style
#' @param fp_p paragraph formatting properties, see [fp_par()].
#' @param fp_t default text formatting properties. This is used as
#' text formatting properties, see [fp_text()]. If NULL (default), the
#' paragraph will used the default text formatting properties (defined by
#' the `base_on` argument).
#' @examples
#' library(officer)
#'
#' doc <- read_docx()
#'
#' doc <- docx_set_paragraph_style(
#'   doc,
#'   style_id = "rightaligned",
#'   style_name = "Explicit label",
#'   fp_p = fp_par(text.align = "right", padding = 20),
#'   fp_t = fp_text_lite(
#'     bold = TRUE,
#'     shading.color = "#FD34F0",
#'     color = "white")
#' )
#'
#' doc <- body_add_par(doc,
#'   value = "This is a test",
#'   style = "Explicit label")
#'
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#' docx_file
docx_set_paragraph_style <- function(x, style_id, style_name, base_on = "Normal", fp_p = fp_par(), fp_t = NULL) {
  styles_file <- file.path(x$package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)

  if (grepl("[^a-zA-Z0-9\\-]+", style_id)) {
    stop("`style_id` should only contain '-', numbers and ascii characters.")
  }

  node_styles <- xml_find_first(doc, "/w:styles")

  fp_p$word_style <- NULL

  if (!is.null(fp_t)){
    fp_t_xml <- rpr_wml(fp_t)
  } else {
    fp_t_xml <- ""
  }
  base_on <- get_style_id(data = x$styles, style = base_on, type = "paragraph")

  xml_code <- paste0(
    sprintf("<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"%s\">", style_id),
    sprintf("<w:name w:val=\"%s\"/>", style_name),
    sprintf("<w:basedOn w:val=\"%s\"/>", base_on),
    ppr_wml(fp_p), fp_t_xml,
    "</w:style>"
  )

  node_style <- xml_child(node_styles, sprintf("w:style[@w:styleId='%s']", style_id))
  if (inherits(node_style, "xml_missing")) {
    xml_add_child(node_styles, as_xml_document(xml_code))
  } else {
    xml_replace(node_style, as_xml_document(xml_code))
  }

  write_xml(doc, file = styles_file)
  styles <- read_docx_styles(x$package_dir)
  x$styles <- styles

  x
}

#' @export
#' @title Add character style in a Word document
#' @description The function lets you add or modify Word character styles.
#' @param x an rdocx object
#' @param style_id a unique style identifier for Word.
#' @param style_name a unique label associated with the style identifier.
#' This label is the name of the style when Word edit the document.
#' @param base_on the character style name used as base style
#' @param fp_t Text formatting properties, see [fp_text()].
#' @examples
#' library(officer)
#' doc <- read_docx()
#'
#' doc <- docx_set_character_style(
#'   doc,
#'   style_id = "newcharstyle",
#'   style_name = "label for char style",
#'   base_on = "Default Paragraph Font",
#'   fp_text_lite(
#'     shading.color = "red",
#'     color = "white")
#' )
#' paragraph <- fpar(
#'   run_wordtext("hello",
#'     style_id = "newcharstyle"))
#'
#' doc <- body_add_fpar(doc, value = paragraph)
#' docx_file <- print(doc, target = tempfile(fileext = ".docx"))
#' docx_file
docx_set_character_style <- function(x, style_id, style_name, base_on, fp_t = fp_text_lite()) {
  styles_file <- file.path(x$package_dir, "word/styles.xml")
  doc <- read_xml(styles_file)
  node_styles <- xml_find_first(doc, "/w:styles")

  if (grepl("[^a-zA-Z0-9\\-]+", style_id)) {
    stop("`style_id` should only contain '-', numbers and ascii characters.")
  }

  base_on <- get_style_id(data = x$styles, style = base_on, type = "character")

  xml_code <- paste0(
    sprintf("<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"character\" w:customStyle=\"1\" w:styleId=\"%s\">", style_id),
    sprintf("<w:name w:val=\"%s\"/>", style_name),
    sprintf("<w:basedOn w:val=\"%s\"/>", base_on),
    rpr_wml(fp_t),
    "</w:style>"
  )

  node_style <- xml_child(node_styles, sprintf("w:style[@w:styleId='%s']", style_id))
  if (inherits(node_style, "xml_missing")) {
    xml_add_child(node_styles, as_xml_document(xml_code))
  } else {
    xml_replace(node_style, as_xml_document(xml_code))
  }

  write_xml(doc, file = styles_file)
  styles <- read_docx_styles(x$package_dir)
  x$styles <- styles

  x
}


#' @export
#' @title Read document properties
#' @description Read Word or PowerPoint document properties
#' and get results in a data.frame.
#' @param x an `rdocx` or `rpptx` object
#' @examples
#' x <- read_docx()
#' doc_properties(x)
#' @return a data.frame
#' @family functions for Word document informations
#' @family functions for reading presentation informations
doc_properties <- function(x) {
  if (inherits(x, "rdocx")) {
    cp <- x$doc_properties
  } else if (inherits(x, "rpptx") || inherits(x, "rxlsx")) {
    cp <- x$core_properties
  } else {
    stop("x should be a rpptx or a rdocx or a rxlsx object.")
  }

  properties_custom <- x$doc_properties_custom

  out_custom <- data.frame(
    tag = properties_custom[, "name"],
    value = properties_custom[, "value"],
    stringsAsFactors = FALSE
  )
  out <- data.frame(
    tag = cp[, "name"],
    value = cp[, "value"],
    stringsAsFactors = FALSE
  )
  out <- rbind(out, out_custom)
  row.names(out) <- NULL
  out
}

#' @export
#' @title Set document properties
#' @description set Word or PowerPoint document properties. These are not visible
#' in the document but are available as metadata of the document.
#'
#' Any character property can be added as a document property.
#' It provides an easy way to insert arbitrary fields. Given the challenges
#' that can be encountered with find-and-replace in word with officer, the
#' use of document fields and quick text fields provides a much more robust
#' approach to automatic document generation from R.
#' @note
#' The "last modified" and "last modified by" fields will be automatically be updated
#' when the file is written.
#' @param x an rdocx or rpptx object
#' @param title,subject,creator,description text fields
#' @param created a date object
#' @param ... named arguments (names are field names), each element is a single
#' character value specifying value associated with the corresponding field name.
#' @param values a named list (names are field names), each element is a single
#' character value specifying value associated with the corresponding field name.
#' If `values` is provided, argument `...` will be ignored.
#' @examples
#' x <- read_docx()
#' x <- set_doc_properties(x, title = "title",
#'   subject = "document subject", creator = "Me me me",
#'   description = "this document is empty",
#'   created = Sys.time(),
#'   yoyo = "yok yok",
#'   glop = "pas glop")
#' x <- doc_properties(x)
#' @family functions for Word document informations
set_doc_properties <- function( x, title = NULL, subject = NULL,
                                creator = NULL, description = NULL, created = NULL,
                                ..., values = NULL){

  if( inherits(x, "rdocx"))
    cp <- x$doc_properties
  else if( inherits(x, "rpptx")) cp <- x$core_properties
  else stop("x should be a rpptx or rdocx object.")

  if( !is.null(title) ) cp['title','value'] <- title
  if( !is.null(subject) ) cp['subject','value'] <- subject
  if( !is.null(creator) ) cp['creator','value'] <- creator
  if( !is.null(description) ) cp['description','value'] <- description
  if( !is.null(created) ) cp['created','value'] <- format( created, "%Y-%m-%dT%H:%M:%SZ")

  if (is.null(values)) {
    values <- list(...)
  }

  if (length(values) > 0) {
    x$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.custom-properties+xml",
               "/docProps/custom.xml")
    )
    x$rel$add(id = paste0("rId", x$rel$get_next_id()),
                type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
                target = "docProps/custom.xml")

    custom_props <- x$doc_properties_custom
    for(i in seq_along(values)) {
      custom_props[names(values)[i], 'value'] <- enc2utf8(values[[i]])
    }
    x$doc_properties_custom <- custom_props
  }

  if( inherits(x, "rdocx"))
    x$doc_properties <- cp
  else x$core_properties <- cp

  x
}


#' @export
#' @title 'Word' page layout
#' @description Get page width, page height and margins (in inches). The return values
#' are those corresponding to the section where the cursor is.
#' @param x an \code{rdocx} object
#' @examples
#' docx_dim(read_docx())
#' @family functions for Word document informations
docx_dim <- function(x){
  cursor <- as.character(x$officer_cursor)
  if (is.na(cursor)) {
    next_section <- xml_find_first(x$doc_obj$get(), "/w:document/w:body/w:sectPr")
  } else {
    xpath_ <- paste0(
      file.path( cursor, "following-sibling::w:sectPr"),
      "|",
      file.path( cursor, "following-sibling::w:p/w:pPr/w:sectPr"),
      "|",
      "//w:sectPr"
    )
    next_section <- xml_find_first(x$doc_obj$get(), xpath_)
  }

  sd <- section_dimensions(next_section)
  sd$page <- sd$page / (20*72)
  sd$margins <- sd$margins / (20*72)
  sd
}


#' @export
#' @title List Word bookmarks
#' @description List bookmarks id that can be found in a
#' 'Word' document.
#' @param x an \code{rdocx} object
#' @examples
#' library(officer)
#'
#' doc_1 <- read_docx()
#' doc_1 <- body_add_par(doc_1, "centered text", style = "centered")
#' doc_1 <- body_bookmark(doc_1, "text_to_replace_1")
#' doc_1 <- body_add_par(doc_1, "centered text", style = "centered")
#' doc_1 <- body_bookmark(doc_1, "text_to_replace_2")
#'
#' docx_bookmarks(doc_1)
#'
#' docx_bookmarks(read_docx())
#' @family functions for Word document informations
docx_bookmarks <- function(x){
  stopifnot(inherits(x, "rdocx"))

  doc_ <- xml_find_all(x$doc_obj$get(), "//w:bookmarkStart[@w:name]")
  setdiff(xml_attr(doc_, "name"), "_GoBack")
}

#' @export
#' @title Replace styles in a 'Word' Document
#' @description Replace styles with others in a 'Word' document. This function
#' can be used for paragraph, run/character and table styles.
#' @param x an rdocx object
#' @param mapstyles a named list, names are the replacement style,
#' content (as a character vector) are the styles to be replaced.
#' Use [styles_info()] to display available styles.
#' @examples
#' # creating a sample docx so that we can illustrate how
#' # to change styles
#' doc_1 <- read_docx()
#'
#' doc_1 <- body_add_par(doc_1, "A title", style = "heading 1")
#' doc_1 <- body_add_par(doc_1, "Another title", style = "heading 2")
#' doc_1 <- body_add_par(doc_1, "Hello world!", style = "Normal")
#' file <- print(doc_1, target = tempfile(fileext = ".docx"))
#'
#' # now we can illustrate how
#' # to change styles with `change_styles`
#' doc_2 <- read_docx(path = file)
#' mapstyles <- list(
#'   "centered" = c("Normal", "heading 2"),
#'   "strong" = "Default Paragraph Font"
#' )
#' doc_2 <- change_styles(doc_2, mapstyles = mapstyles)
#' print(doc_2, target = tempfile(fileext = ".docx"))
change_styles <- function( x, mapstyles ){

  if( is.null(mapstyles) || length(mapstyles) < 1 ) return(x)

  table_styles <- styles_info(x, type = c("paragraph", "character", "table"))

  from_styles <- unique( as.character( unlist(mapstyles) ) )
  to_styles <- unique( names( mapstyles) )

  if( any( is.na( mfrom <- match( from_styles, table_styles$style_name ) ) ) ){
    stop("could not find style ", paste0( shQuote(from_styles[is.na(mfrom)]), collapse = ", " ), ".", call. = FALSE)
  }
  if( any( is.na( mto <- match( to_styles, table_styles$style_name ) ) ) ){
    stop("could not find style ", paste0( shQuote(to_styles[is.na(mto)]), collapse = ", " ), ".", call. = FALSE)
  }

  mapping <- mapply(function(from, to) {
    id_to <- which( table_styles$style_name %in% to )
    id_to <- table_styles$style_id[id_to]

    id_from <- which( table_styles$style_name %in% from )
    types <- substring(table_styles$style_type[id_from], first = 1, last = 1)
    types[types %in% "c"] <- "r"
    types[types %in% "t"] <- "tbl"
    id_from <- table_styles$style_id[id_from]

    data.frame( from = id_from, to = rep(id_to, length(from)), types = types, stringsAsFactors = FALSE )
  }, mapstyles, names(mapstyles), SIMPLIFY = FALSE)

  mapping <- do.call(rbind, mapping)
  row.names(mapping) <- NULL

  for(i in seq_len( nrow(mapping) )){
    all_nodes <- xml_find_all(x$doc_obj$get(), sprintf("//w:%sStyle[@w:val='%s']", mapping$types[i], mapping$from[i]))
    xml_attr(all_nodes, "w:val") <- rep(mapping$to[i], length(all_nodes) )
  }

  x
}



#' @export
#' @title Body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_xml(doc)
#' @keywords internal
docx_body_xml <- function( x ){
  x$doc_obj$get()
}
#' @export
#' @title xml element on which cursor is
#' @description Get the current block element as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer. If the
#' document is empty, this block will be set to NULL.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_current_block_xml(doc)
#' @keywords internal
docx_current_block_xml <- function( x ){
  ooxml_on_cursor(x$officer_cursor, x$doc_obj$get())
}

#' @export
#' @title Body xml document
#' @description Get the body document as xml. This function
#' is not to be used by end users, it has been implemented
#' to allow other packages to work with officer.
#' @param x an rdocx object
#' @examples
#' doc <- read_docx()
#' docx_body_relationship(doc)
#' @keywords internal
docx_body_relationship <- function( x ){
  x$doc_obj$relationship()
}

