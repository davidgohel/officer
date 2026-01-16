read_core_properties <- function(package_dir) {
  filename <- file.path(package_dir, "docProps/core.xml")
  if (!file.exists(filename)) {
    # docx and pptx files exported from Google Docs or Slides do not have properties set. Use one from the template to get by
    warning(
      "No properties found. Using properties from template file in officer."
    )
    filename <- system.file(package = "officer", "template/core.xml")
  }
  doc <- read_xml(filename)
  ns_ <- xml_ns(doc)

  all_ <- xml_find_all(doc, "/cp:coreProperties/*")

  if (length(all_) < 1) {
    out <- list(
      data = structure(
        character(0),
        .Dim = c(0L, 4L),
        .Dimnames = list(
          NULL,
          c("ns", "name", "attrs", "value")
        )
      ),
      ns = c(
        cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        dc = "http://purl.org/dc/elements/1.1/",
        dcmitype = "http://purl.org/dc/dcmitype/",
        dcterms = "http://purl.org/dc/terms/",
        xsi = "http://www.w3.org/2001/XMLSchema-instance"
      )
    )
    return(out)
  }

  names_ <- sapply(all_, xml_name, ns = ns_)
  names_ <- strsplit(names_, ":")

  # concat all attrs in single chars
  attrs <- sapply(xml_attrs(all_), function(ns_) {
    paste0('xsi:', names(ns_), '=\"', ns_, '\"', collapse = " ")
  })
  attrs <- ifelse(sapply(xml_attrs(all_), length) < 1, "", paste0(" ", attrs))

  propnames <- sapply(names_, function(x) x[2])

  props <- matrix(
    c(sapply(names_, function(x) x[1]), propnames, attrs, xml_text(all_)),
    ncol = 4
  )
  colnames(props) <- c("ns", "name", "attrs", "value")
  rownames(props) <- propnames
  z <- list(data = props, ns = unclass(ns_))
  class(z) <- "core_properties"
  z
}

default_cp_tags <- list(
  title = list(ns = "dc", attrs = NA_character_),
  subject = list(ns = "dc", attrs = NA_character_),
  creator = list(ns = "dc", attrs = NA_character_),
  keywords = list(ns = "cp", attrs = NA_character_),
  description = list(ns = "dc", attrs = NA_character_),
  lastModifiedBy = list(ns = "cp", attrs = NA_character_),
  revision = list(ns = "cp", attrs = NA_character_),
  created = list(ns = "dcterms", attrs = 'xsi:type="dcterms:W3CDTF"'),
  modified = list(ns = "dcterms", attrs = 'xsi:type="dcterms:W3CDTF"'),
  category = list(ns = "cp", attrs = NA_character_)
)


#' @export
`[<-.core_properties` <- function(x, i, j, value) {
  if (!i %in% x$data[, "name"]) {
    attrs <- ifelse(
      is.na(default_cp_tags[[i]]$attrs),
      "",
      paste0(" ", default_cp_tags[[i]]$attrs)
    )
    new <- matrix(
      c(default_cp_tags[[i]]$ns, i, attrs, value),
      nrow = 1,
      dimnames = list(i, colnames(x$data))
    )
    x$data <- rbind(x$data, new)
  } else {
    x$data[i, j] <- value
  }
  x
}


#' @export
`[.core_properties` <- function(x, i, j) {
  x$data[i, j]
}


write_core_properties <- function(core_matrix, package_dir) {
  ns_ <- core_matrix$ns
  core_matrix <- core_matrix$data
  if (!is.matrix(core_matrix)) {
    stop("core_properties should be stored in a character matrix.")
  }

  ns_ <- paste0('xmlns:', names(ns_), '=\"', ns_, '\"', collapse = " ")
  xml_ <- paste0(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '\n',
    '<cp:coreProperties ',
    ns_,
    '>'
  )
  properties <- sprintf(
    "<%s:%s%s>%s</%s:%s>",
    core_matrix[, "ns"],
    core_matrix[, "name"],
    core_matrix[, "attrs"],
    htmlEscapeCopy(core_matrix[, "value"]),
    core_matrix[, "ns"],
    core_matrix[, "name"]
  )
  xml_ <- paste0(
    xml_,
    paste0(properties, collapse = ""),
    "</cp:coreProperties>"
  )
  props_dir = file.path(package_dir, "docProps")
  dir.create(props_dir, recursive = TRUE, showWarnings = FALSE)
  filename <- file.path(props_dir, "core.xml")
  writeLines(enc2utf8(xml_), filename, useBytes = TRUE)
  invisible()
}

drop_templatenode_from_app <- function(package_dir) {
  file <- file.path(package_dir, "docProps", "app.xml")
  if (!file.exists(file)) {
    return()
  }
  doc <- read_xml(file)
  node_template <- xml_child(doc, "d1:Template")
  if (!inherits(node_template, "xml_missing")) {
    xml_remove(node_template)
    write_xml(doc, file)
  }
}


# app.xml properties ----

read_app_properties <- function(package_dir) {
  file <- file.path(package_dir, "docProps", "app.xml")
  if (!file.exists(file)) {
    return(NULL)
  }
  doc <- read_xml(file)
  ns <- xml_ns(doc)

  hyperlink_base_node <- xml_find_first(doc, "//d1:HyperlinkBase", ns = ns)
  hyperlink_base <- if (inherits(hyperlink_base_node, "xml_missing")) {
    NA_character_
  } else {
    xml_text(hyperlink_base_node)
  }

  company_node <- xml_find_first(doc, "//d1:Company", ns = ns)
  company <- if (inherits(company_node, "xml_missing")) {
    NA_character_
  } else {
    xml_text(company_node)
  }

  z <- list(
    data = data.frame(
      name = c("HyperlinkBase", "Company"),
      value = c(hyperlink_base, company),
      stringsAsFactors = FALSE
    ),
    file = file
  )
  class(z) <- "app_properties"
  z
}


#' @export
`[<-.app_properties` <- function(x, i, j, value) {
  if (!i %in% x$data$name) {
    new_row <- data.frame(name = i, value = value, stringsAsFactors = FALSE)
    x$data <- rbind(x$data, new_row)
  } else {
    x$data[x$data$name == i, j] <- value
  }
  x
}


#' @export
`[.app_properties` <- function(x, i, j) {
  x$data[x$data$name == i, j]
}

#' @importFrom xml2 xml_set_text
write_app_properties <- function(app_props, package_dir) {
  if (is.null(app_props)) {
    return(invisible())
  }

  file <- file.path(package_dir, "docProps", "app.xml")
  if (!file.exists(file)) {
    return(invisible())
  }

  doc <- read_xml(file)
  ns <- xml_ns(doc)

  for (i in seq_len(nrow(app_props$data))) {
    prop_name <- app_props$data$name[i]
    prop_value <- app_props$data$value[i]

    if (is.na(prop_value) || prop_value == "") {
      next
    }

    xpath <- paste0("//d1:", prop_name)
    node <- xml_find_first(doc, xpath, ns = ns)
    if (inherits(node, "xml_missing")) {
      new_node <- read_xml(sprintf(
        "<%s>%s</%s>",
        prop_name,
        htmlEscapeCopy(prop_value),
        prop_name
      ))
      xml_add_child(doc, new_node)
    } else {
      xml_set_text(node, prop_value)
    }
  }

  write_xml(doc, file)
  invisible()
}
