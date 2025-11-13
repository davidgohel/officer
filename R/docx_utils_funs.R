is_string <- function(x) {
  is.character(x) && length(x) == 1 && !is.na(x)
}
is_scalar_datetime <- function(x) {
  inherits(x, "POSIXt") && length(x) == 1 && !is.na(x)
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
length.rdocx <- function(x) {
  xml_length(xml_child(x$doc_obj$get(), "w:body"))
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
  values = NULL
) {
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
      if (
        !is.null(values[[i]]) &&
          length(values[[i]]) == 1 &&
          !is.na(values[[i]])
      ) {
        .value. <- format(values[[i]])
        .value. <- enc2utf8(.value.)
      } else if (
        !is.null(values[[i]]) &&
          length(values[[i]]) == 1 &&
          is.na(values[[i]])
      ) {
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
docx_bookmarks <- function(x) {
  stopifnot(inherits(x, "rdocx"))

  doc_ <- xml_find_all(x$doc_obj$get(), "//w:bookmarkStart[@w:name]")
  setdiff(xml_attr(doc_, "name"), "_GoBack")
}

#' @export
#' @title 'Word' page layout
#' @description Get page width, page height and margins (in inches). The return values
#' are those corresponding to the section where the cursor is.
#' @param x an `rdocx` object
#' @examples
#' docx_dim(read_docx())
#' @family functions for Word document informations
docx_dim <- function(x) {
  cursor <- as.character(x$officer_cursor)
  if (is.na(cursor)) {
    next_section <- xml_find_first(
      x$doc_obj$get(),
      "/w:document/w:body/w:sectPr"
    )
  } else {
    xpath_ <- paste0(
      file.path(cursor, "following-sibling::w:sectPr"),
      "|",
      file.path(cursor, "following-sibling::w:p/w:pPr/w:sectPr"),
      "|",
      "//w:sectPr"
    )
    next_section <- xml_find_first(x$doc_obj$get(), xpath_)
  }

  sd <- section_dimensions(next_section)
  sd$page <- sd$page / (20 * 72)
  sd$margins <- sd$margins / (20 * 72)
  sd
}
