#' @export
#' @title Create a 'Word' document object
#' @description `read_docx()` is the starting point for creating Word documents
#' from R. It creates an R object representing a Word document that can be
#' manipulated programmatically. When called without arguments, it creates an
#' empty document based on a default template. When provided with a path, it
#' reads an existing Word document (`.docx`) or template (`.dotx`) file.
#'
#' Once created, you can:
#' - Add content from R: Insert text, formatted paragraphs ([fpar()]),
#'   tables ([body_add_table()]), plots ([body_add_plot()]), images,
#'   page breaks, table of contents, and more using the `body_add_*()` family
#'   of functions.
#' - Read and inspect content: Use [docx_summary()] to extract and analyze
#'   the document's content, structure, and formatting as a data frame.
#' - Write to file: Save the document to disk using
#'   `print(x, target = "path/to/file.docx")`.
#'
#' @param path path to the docx file to use as base document.
#' `dotx` file are supported.
#' @return an object of class `rdocx`.
#' @section styles:
#'
#' The template file (specified via `path` or the default template) determines
#' the available paragraph styles, character styles, and table styles in your
#' document. These styles control the appearance of headings, body text, tables,
#' and other elements.
#'
#' When you use functions like `body_add_par(style = "heading 2")`, the style
#' name must exist in the template. You can:
#' - Use [styles_info()] to list all available styles in your document
#' - Create a custom template in Word with your organization's styles and branding
#' - Use the default template for standard documents
#'
#' The document layout (page size, margins, headers and footer content,
#' orientation) also comes from the
#' template and can be modified using [body_set_default_section()] or by
#' adding section breaks with [body_end_section_continuous()] and related functions.
#' @examples
#' library(officer)
#'
#' # This example demonstrates how to create
#' # an empty document -----
#'
#' ## Create a new Word document
#' doc <- read_docx()
#' ## Save the document
#' output_file <- print(doc, target = tempfile(fileext = ".docx"))
#'
#' @example inst/examples/example_read_docx_1.R
#' @example inst/examples/example_read_docx_2.R
#' @seealso Save a 'Word' document object to a file with [print.rdocx()],
#' add content with functions [body_add_par()], [body_add_plot()],
#' [body_add_table()], change settings with [docx_set_settings()], set
#' properties with [set_doc_properties()],  read 'Word' styles with
#' [styles_info()].
#' @section Illustrations:
#'
#' \if{html}{\figure{read_docx.png}{options: width=80\%}}
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

  obj$settings <- update_docx_settings_from_file(
    x = docx_settings(),
    file = file.path(package_dir, "word", "settings.xml")
  )

  obj$rel <- relationship$new()
  obj$rel$feed_from_xml(file.path(package_dir, "_rels", ".rels"))

  obj$doc_properties_custom <- read_custom_properties(package_dir)
  obj$doc_properties <- read_core_properties(package_dir)
  obj$content_type <- content_type$new(package_dir)
  obj$doc_obj <- body_part$new(
    package_dir,
    main_file = "document.xml",
    cursor = "/w:document/w:body/*[1]",
    body_xpath = "/w:document/w:body"
  )
  obj$styles <- read_docx_styles(package_dir)
  obj$officer_cursor <- officer_cursor(obj$doc_obj$get())

  obj$headers <- update_hf_list(part_list = list(), type = "header", package_dir = package_dir)
  obj$footers <- update_hf_list(part_list = list(), type = "footer", package_dir = package_dir)

  if (!file.exists(file.path(package_dir, "word", "comments.xml"))) {
    file.copy(
      system.file(package = "officer", "template", "comments.xml"),
      file.path(package_dir, "word", "comments.xml"),
      copy.mode = FALSE
    )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", "/word/comments.xml")
    )
  }

  obj$comments <- docx_part$new(
    package_dir,
    main_file = "comments.xml",
    cursor = "/w:comments/*[last()]", body_xpath = "/w:comments"
  )

  if (!file.exists(file.path(package_dir, "word", "footnotes.xml"))) {
    file.copy(
      system.file(package = "officer", "template", "footnotes.xml"),
      file.path(package_dir, "word", "footnotes.xml"),
      copy.mode = FALSE
    )
    obj$content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "/word/footnotes.xml")
    )
  }

  obj$footnotes <- footnotes_part$new(
    package_dir,
    main_file = "footnotes.xml",
    cursor = "/w:footnotes/*[last()]",
    body_xpath = "/w:footnotes"
  )

  default_refs <- obj$styles[obj$styles$is_default, ]
  obj$default_styles <- setNames(as.list(default_refs$style_name), default_refs$style_type)

  last_sect <- xml_find_first(obj$doc_obj$get(), "/w:document/w:body/w:sectPr[last()]")
  obj$sect_dim <- section_dimensions(last_sect)

  obj <- cursor_end(obj)
  obj
}


#' @export
#' @title Number of blocks inside an rdocx object
#' @description Return the number of blocks inside an rdocx object.
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
#' @description Read Word styles and get results in
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
#' @family functions for reading presentation information
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

is_string <- function(x) {
  is.character(x) && length(x) == 1 && !is.na(x)
}
is_scalar_datetime <- function(x) {
  inherits(x, "POSIXt") && length(x) == 1 && !is.na(x)
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
#' These pairs of *key-value* are added as custom properties. If a value is
#' `NULL` or `NA`, the corresponding field is set to '' in the document properties.
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
set_doc_properties <- function(
    x,
    title = NULL,
    subject = NULL,
    creator = NULL,
    description = NULL,
    created = NULL,
    ...,
    values = NULL) {
  if (inherits(x, "rdocx")) {
    cp <- x$doc_properties
  } else if (inherits(x, "rpptx")) {
    cp <- x$core_properties
  } else {
    stop("x should be a rpptx or rdocx object.")
  }

  if (!is.null(title)) {
    if (!is_string(title)) {
      cli::cli_warn(
        c(
          "!" = "The value for property 'title' is not a string.",
          "i" = "It will be set to '' in the document properties"
        )
      )
      title <- ""
    }
    cp["title", "value"] <- title
  }

  if (!is.null(subject)) {
    if (!is_string(subject)) {
      cli::cli_warn(
        c(
          "!" = "The value for property 'subject' is not a string.",
          "i" = "It will be set to '' in the document properties"
        )
      )
      subject <- ""
    }
    cp["subject", "value"] <- subject
  }

  if (!is.null(creator)) {
    if (!is_string(creator)) {
      cli::cli_warn(
        c(
          "!" = "The value for property 'creator' is not a string.",
          "i" = "It will be set to '' in the document properties"
        )
      )
      creator <- ""
    }
    cp["creator", "value"] <- creator
  }

  if (!is.null(description)) {
    if (!is_string(description)) {
      cli::cli_warn(
        c(
          "!" = "The value for property 'description' is not a string.",
          "i" = "It will be set to '' in the document properties"
        )
      )
      description <- ""
    }
    cp["description", "value"] <- description
  }

  if (!is.null(created)) {
    if (!is_scalar_datetime(created)) {
      cli::cli_warn(
        c(
          "!" = "The value for property 'created' is not a date-time.",
          "i" = "It will not be set in the document properties"
        )
      )
    } else {
      cp["created", "value"] <- format(created, "%Y-%m-%dT%H:%M:%SZ")
    }
  }

  if (is.null(values)) {
    values <- list(...)
  }

  if (length(values) > 0) {
    x$content_type$add_override(
      setNames(
        "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        "/docProps/custom.xml"
      )
    )
    x$rel$add(
      id = paste0("rId", x$rel$get_next_id()),
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
      target = "docProps/custom.xml"
    )

    custom_props <- x$doc_properties_custom
    for (i in seq_along(values)) {
      .value. <- ""
      if (!is.null(values[[i]]) &&
        length(values[[i]]) == 1 &&
        !is.na(values[[i]])) {
        .value. <- format(values[[i]])
        .value. <- enc2utf8(.value.)
      } else if (!is.null(values[[i]]) &&
        length(values[[i]]) == 1 &&
        is.na(values[[i]])) {
        .value. <- ""
      } else if (is.null(values[[i]])) {
        .value. <- ""
      } else if (length(values[[i]]) != 1) {
        .value. <- ""
        cli::cli_warn(
          c(
            "!" = "The value for property '{names(values)[i]}' is not a single character value.",
            "i" = "It will be set to '' in the document properties"
          )
        )
      }
      custom_props[names(values)[i], "value"] <- .value.
    }
    x$doc_properties_custom <- custom_props
  }

  if (inherits(x, "rdocx")) {
    x$doc_properties <- cp
  } else {
    x$core_properties <- cp
  }

  x
}


#' @export
#' @title 'Word' page layout
#' @description Get page width, page height and margins (in inches). The return values
#' are those corresponding to the section where the cursor is.
#' @param x an `rdocx` object
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
#' @param x an `rdocx` object
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
