xml_document_to_chrs <- function(xml_doc) {
  # Write the XML document to a temporary file
  tf <- tempfile(fileext = ".xml")
  write_xml(xml_doc, options = c("format"), file = tf)
  # Read the XML file as a character vector
  xml_str <- readLines(tf, warn = FALSE)
  xml_str <- gsub("^[[:blank:]]+", "", xml_str)
  # Prepend newlines to the XML string
  # where the opening tags are found
  m <- gregexpr("(?=<[^/])", xml_str, perl = TRUE)
  regmatches(xml_str, m) <- "\n"
  # The following can be replaced with gregexpr()
  # but there is minor performance gain with the following
  writeLines(xml_str, tf, useBytes = TRUE)
  xml_str <- readLines(tf, warn = FALSE)

  xml_str <- xml_str[xml_str != ""]
  xml_str
}

# replace body xml object with a new xml stored as a character vector
replace_xml_body_from_chr <- function(x, xml_str) {
  tf_xml <- tempfile(fileext = ".txt")
  writeLines(xml_str, tf_xml, useBytes = TRUE)
  x$doc_obj$replace_xml(tf_xml)
  x
}

#' @title Fix empty IDs in WML string
#' @description
#' This function corrects empty IDs found in a WordprocessingML (WML) string.
#' It ensures that all IDs are unique.
#' @param xml_str A character string containing WML with potentially empty embed IDs.
#' @return A character string (WML) where all empty IDs have been replaced
#' with unique IDs.
#' @noRd
fix_empty_ids_in_wml <- function(xml_str) {
  m <- regexpr(" id=\"\"", xml_str)
  id_values <- seq_len(sum(m > 0))
  regmatches(xml_str, m) <- sprintf(" id=\"%.0f\"", id_values)
  xml_str
}

#' @title Register images in relationships part of the Word package
#' @description
#' Ensures that each image referenced in the document is correctly declared in
#' the .rels file and assigned a unique relationship ID. Copies images if needed.
#' @param path_table data.frame containing a column named `path` that define the
#' image paths to be registered.
#' @param relationships A relationships object where the images will be registered.
#' @param package_dir The root directory of the unzipped Word package (used to store media).
#' @param media_dir The path within the package where image files are stored (default: "word/media").
#' @param media_rel_dir Relative path from the relationships part to the image (default: "media").
#' @noRd
register_images_in_relationships <- function(
  path_table,
  relationships,
  package_dir,
  media_dir = "word/media",
  media_rel_dir = "media"
) {
  path_table$id <- rep(NA_character_, nrow(path_table))
  for (i in seq_len(nrow(path_table))) {
    path <- path_table[[i, "path"]]
    dest_basename <- fake_newname(path)
    img_path <- file.path(package_dir, media_dir)
    if (!file.exists(file.path(img_path, dest_basename))) {
      dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
      file.copy(from = path, to = file.path(img_path, dest_basename))
    }
    if (
      !file.path(media_rel_dir, dest_basename) %in%
        relationships$get_data()$target
    ) {
      rid <- sprintf("rId%.0f", relationships$get_next_id())
      relationships$add(
        id = rid,
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target = file.path(media_rel_dir, dest_basename)
      )
    } else {
      reldf <- relationships$get_data()
      rid <- reldf$id[basename(reldf$target) %in% dest_basename]
    }
    path_table$id[i] <- rid
  }
  path_table
}

#' @title Fix <a:blip> image references in WML
#' @description
#' This function replaces custom `<a:blip>` tags used by the officer package
#' with standard `<a:blip>` XML nodes, enabling proper image embedding
#' in a WordprocessingML (WML) string. Each embedded image is registered
#' in the relationships file using a unique `rId`. If this function is not executed,
#' the resulting Word document may be corrupted or display missing images.
#' @param xml_str A character string containing WML (WordprocessingML XML) where <a:blip> tags must be replaced.
#' @param doc_obj An `xml_document` object corresponding to the part being processed (usually the main document or a header/footer).
#' @param relationships A `relationship` object used to track and register media references.
#' @param package_dir The root directory of the unzipped Word package (used to store media files).
#' @param media_dir Path within the package where image files are stored. Defaults to `"word/media"`.
#' @param media_rel_dir Relative path from the `.rels` file to the image folder. Defaults to `"media"`.
#' @return A character string (WML) with all <a:blip> nodes replaced by valid <a:blip> image references using registered `rId`s.
#' @noRd
fix_img_refs_in_wml <- function(
  xml_str,
  doc_obj,
  relationships,
  package_dir,
  media_dir = "word/media",
  media_rel_dir = "media"
) {
  has_match <- grepl("<a:blip r:embed=\"[^\"]+\"", xml_str) &
    !grepl("<a:blip r:embed=\"rId[0-9]+\"", xml_str)

  if (!any(has_match)) {
    return(xml_str)
  }

  img_nodes_chr <- xml_str[has_match]

  path_values <- gsub(
    "<a:blip r:embed=\"([^\"]+)\".*",
    "\\1",
    img_nodes_chr
  )
  path_values <- unique(path_values)

  path_table <- data.frame(
    path = path_values,
    stringsAsFactors = FALSE
  )

  path_table <- register_images_in_relationships(
    path_table = path_table,
    relationships = relationships,
    package_dir = package_dir,
    media_dir = media_dir,
    media_rel_dir = media_rel_dir
  )
  for (i in seq_len(nrow(path_table))) {
    path <- path_table[[i, "path"]]
    rid <- path_table[[i, "id"]]
    xml_str <- gsub(
      sprintf("<a:blip r:embed=\"%s\"", path),
      sprintf("<a:blip r:embed=\"%s\"", rid),
      xml_str,
      fixed = TRUE
    )
  }
  xml_str
}


#' @title Fix <asvg:svgBlip> image references in WML
#' @description
#' This function replaces custom `<asvg:svgBlip>` tags used by the officer package
#' with standard `<asvg:svgBlip>` XML nodes, enabling proper SVG embedding
#' in a WordprocessingML (WML) string. Each embedded SVG is registered
#' in the relationships file using a unique `rId`. If this function is not executed,
#' the resulting Word document may be corrupted or display missing images.
#'
#' The function also ensures that fallback images are used when SVG rendering
#' is not supported by the target application (e.g., Word for Windows), again,
#' this is about setting the correct unique `rId`.
#' @param xml_str A character string containing WML (WordprocessingML XML) where <p:svg> tags must be replaced.
#' @param doc_obj An `xml_document` object corresponding to the part being processed (usually the main document or a header/footer).
#' @param relationships A `relationship` object used to track and register media references.
#' @param package_dir The root directory of the unzipped Word package (used to store media files).
#' @param media_dir Path within the package where image files are stored. Defaults to `"word/media"`.
#' @param media_rel_dir Relative path from the `.rels` file to the image folder. Defaults to `"media"`.
#' @return A character string (WML) with all <asvg:svgBlip> nodes replaced by valid <asvg:svgBlip> image references using registered `rId`s.
#' @noRd
fix_svg_refs_in_wml <- function(
  xml_str,
  doc_obj,
  relationships,
  package_dir,
  media_dir = "word/media",
  media_rel_dir = "media"
) {
  has_match <- grepl(
    "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"[^\"]+\"/>",
    xml_str
  ) &
    !grepl(
      "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId[0-9]+\"/>",
      xml_str
    )
  if (!any(has_match)) {
    return(xml_str)
  }

  img_nodes_chr <- xml_str[has_match]
  svg_path_values <- gsub(
    "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"([^\"]+)\"/>",
    "\\1",
    img_nodes_chr
  )
  svg_path_values <- unique(svg_path_values)

  path_table <- data.frame(
    path = svg_path_values,
    stringsAsFactors = FALSE
  )

  path_table <- register_images_in_relationships(
    path_table = path_table,
    relationships = relationships,
    package_dir = package_dir,
    media_dir = media_dir,
    media_rel_dir = media_rel_dir
  )
  for (i in seq_len(nrow(path_table))) {
    path <- path_table[[i, "path"]]
    rid <- path_table[[i, "id"]]

    xml_str <- gsub(
      sprintf(
        "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"%s\"/>",
        path
      ),
      sprintf(
        "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"%s\"/>",
        rid
      ),
      xml_str,
      fixed = TRUE
    )
  }

  has_match <- grepl("<a:blip r:embed=\"[^\"]+\">", xml_str) &
    !grepl("<a:blip r:embed=\"rId[0-9]+\">", xml_str)
  img_nodes_chr <- xml_str[has_match]
  img_path_values <- gsub(
    "<a:blip r:embed=\"([^\"]+)\">",
    "\\1",
    img_nodes_chr
  )
  img_path_values <- unique(img_path_values)
  path_table <- data.frame(
    path = img_path_values,
    stringsAsFactors = FALSE
  )
  path_table <- register_images_in_relationships(
    path_table = path_table,
    relationships = relationships,
    package_dir = package_dir,
    media_dir = media_dir,
    media_rel_dir = media_rel_dir
  )
  for (i in seq_len(nrow(path_table))) {
    path <- path_table[[i, "path"]]
    rid <- path_table[[i, "id"]]
    xml_str <- gsub(
      sprintf("<a:blip r:embed=\"%s\">", path),
      sprintf("<a:blip r:embed=\"%s\">", rid),
      xml_str,
      fixed = TRUE
    )
  }
  xml_str
}

#' @title Fix <w:hyperlink> hyperlink references in WML
#' @description
#' This function replaces custom `<w:hyperlink>` tags (used internally by the officer package)
#' with standard OOXML `<w:hyperlink>` elements in a WordprocessingML (WML) string.
#' It ensures that all external hyperlinks are registered in the relationships file
#' with valid `rId` identifiers and updates the WML string accordingly.
#'
#' If this function is not executed, the resulting Word document may be corrupted
#' or contain invalid or non-functional hyperlinks.
#' @param xml_str A character string containing WML (WordprocessingML XML) with <w:hyperlink> tags to replace.
#' @param relationships A `relationship` object used to register each hyperlink target and assign unique `rId`s.
#' @return A character string (WML) with all <w:hyperlink> nodes replaced by valid <w:hyperlink> references using registered `rId`s.
#' @noRd
fix_hyperlink_refs_in_wml <- function(xml_str, doc_obj) {
  rel <- doc_obj$relationship()
  next_rel_id <- rel$get_next_id()

  has_match <- grepl("<w:hyperlink r:id=\"[^\"]+\"", xml_str) &
    !grepl("<w:hyperlink r:id=\"rId[0-9]+\"", xml_str)
  if (!any(has_match)) {
    return(xml_str)
  }

  hyperlink_nodes_chr <- xml_str[has_match]

  hyperlink_values <- gsub(
    ".*<w:hyperlink\\s+r:id=\"([^\"]+)\"[^>]*>.*", "\\1",
    hyperlink_nodes_chr
  )
  hyperlink_values <- unique(hyperlink_values)
  hyperlinks_table <- data.frame(
    url = hyperlink_values,
    id = seq(from = next_rel_id, along.with = hyperlink_values, by = 1L),
    stringsAsFactors = FALSE
  )

  for (i in seq_len(nrow(hyperlinks_table))) {
    url <- hyperlinks_table[[i, "url"]]
    rid <- sprintf("rId%.0f", hyperlinks_table[[i, "id"]])
    rel$add(
      id = rid,
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = officer_url_decode(url),
      target_mode = "External"
    )
    xml_str <- gsub(
      sprintf("<w:hyperlink r:id=\"%s\"", url),
      sprintf("<w:hyperlink r:id=\"%s\"", rid),
      xml_str,
      fixed = TRUE
    )
  }
  xml_str
}

#' @title Convert custom style tags to OOXML style references
#' @description
#' This function parses a WordprocessingML (WML) string containing officer-specific
#' style tags and replaces them with standard OOXML-compliant `<w:pStyle>` nodes.
#'
#' This conversion is required for Word to correctly interpret and apply the styles
#' defined in the document's style part. Without this processing step, custom styles
#' will not be rendered or recognized.
#' @param xml_str A character string containing WML (WordprocessingML XML) with
#' officer custom style tags.
#' @param styles style table returned by `styles_info()`
#' @return A character string (WML) with all custom style tags replaced by standard OOXML style references.
#' @noRd
convert_custom_styles_in_wml <- function(xml_str, styles) {
  m <- gregexpr("w:(pstlname|tstlname)=\"([^\"]+)\"", xml_str)
  stylenames <- unique(unlist(regmatches(xml_str, m)))

  if (length(stylenames) < 1) {
    return(xml_str)
  }

  stylenames_types <- gsub(
    "w:(pstlname|tstlname)=\"([^\"]+)\"",
    "\\1",
    stylenames
  )
  stylenames <- gsub("w:(pstlname|tstlname)=\"([^\"]+)\"", "\\2", stylenames)

  if (!all(stylenames %in% styles$style_name)) {
    missing_styles <- paste0(
      shQuote(unique(setdiff(stylenames, styles$style_name))),
      collapse = ", "
    )
    stop("Some styles can not be found in the document: ", missing_styles)
  }

  styleids <- styles$style_id[match(stylenames, styles$style_name)]
  for (i in seq_along(stylenames)) {
    if ("pstlname" %in% stylenames_types[i]) {
      xml_str <- gsub(
        sprintf("w:pstlname=\"%s\"", stylenames[i]),
        sprintf("w:val=\"%s\"", styleids[i]),
        xml_str,
        fixed = TRUE
      )
    } else if ("tstlname" %in% stylenames_types[i]) {
      xml_str <- gsub(
        sprintf("w:tstlname=\"%s\"", stylenames[i]),
        sprintf("w:val=\"%s\"", styleids[i]),
        xml_str,
        fixed = TRUE
      )
    }
  }
  xml_str
}


#' @title Process footnote content written by officer
#' @description
#' This function processes a WordprocessingML (WML) string containing footnote content
#' written by officer. It ensures that all custom tags are converted into valid
#' OOXML equivalents by applying the appropriate fix/convert functions, i.e.
#' by copying the content of the footnote into the footnotes part of the document
#' and by setting the correct relationship ID (rId) for each footnote.
#'
#' This step is essential for ensuring that footnote content is compatible with
#' Word's rendering engine and avoids corruption of the document structure.
#'
#' @param x A character string containing WML representing footnote content.
#' @param footnotes An `xml_document` object representing the document footnotes being processed.
#' @param xml_str A character string containing WML (WordprocessingML XML) with
#' officer custom footnotes tags.
#' @return A character string (WML) converted into valid OOXML elements.
#' @noRd
process_footnotes_content <- function(x, footnotes, xml_str) {
  index <- length(xml_find_all(footnotes$get(), "w:footnote")) - 1

  footnotes_summary <- extract_footnotes(x, xml_str, index = index)
  if (nrow(footnotes_summary) < 1) {
    return(xml_str)
  }

  xml_str[footnotes_summary$reference_position] <- sprintf(
    xml_str[footnotes_summary$reference_position],
    footnotes_summary[["rid"]]
  )

  footnotes_summary$content_section <- sprintf(
    footnotes_summary$content_section,
    footnotes_summary$rid
  )

  relationships <- x$relationship()
  if (nrow(footnotes_summary) > 0) {
    next_id <- relationships$get_next_id()

    relationships$add(
      paste0("rId", next_id),
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
      target = "footnotes.xml"
    )
  }
  for (i in seq_len(nrow(footnotes_summary))) {
    rid <- footnotes_summary[[i, "rid"]]
    content_section <- footnotes_summary[[i, "content_section"]]

    newfootnote <- as_xml_document(content_section)
    xml_add_child(footnotes$get(), newfootnote)
  }

  # drop added sections contents from document xml
  lines_index_to_drop <- mapply(
    function(start, stop, str) {
      start:stop
    },
    start = footnotes_summary$start,
    stop = footnotes_summary$stop,
    SIMPLIFY = FALSE
  )
  lines_index_to_drop <- unlist(lines_index_to_drop)
  if (length(lines_index_to_drop) > 0) {
    xml_str <- xml_str[-lines_index_to_drop]
  }

  xml_str
}

#' @title Process comments content written by officer
#' @description
#' This function processes a WordprocessingML (WML) string containing comments content
#' written by officer. It ensures that all custom tags are converted into valid
#' OOXML equivalents by applying the appropriate fix/convert functions, i.e.
#' by copying the content of the comment into the comments part of the document
#' and by setting the correct relationship ID (rId) for each comment.
#'
#' This step is essential for ensuring that comments content is compatible with
#' Word's rendering engine and avoids corruption of the document structure.
#'
#' @param x A character string containing WML representing footnote content.
#' @param comments An `xml_document` object representing the document comments being processed.
#' @param xml_str A character string containing WML (WordprocessingML XML) with
#' officer custom comments tags.
#' @return A character string (WML) converted into valid OOXML elements.
#' @noRd
process_comments_content <- function(x, comments, xml_str) {
  index <- length(xml_find_all(comments$get(), "w:comment"))
  comments_summary <- extract_comments(x, xml_str, index)
  if (nrow(comments_summary) < 1) {
    return(xml_str)
  }
  relationships <- x$relationship()
  if (nrow(comments_summary) > 0) {
    next_id <- relationships$get_next_id()
    relationships$add(
      paste0("rId", next_id),
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
      target = "comments.xml"
    )
  }

  comments_summary$fake_id <- gsub(
    "<w:commentReference w:officer=\"true\" w:id=\"([^\"]+)\">",
    "\\1",
    xml_str[comments_summary$reference_position],
  )
  comments_summary$new_id <- seq(
    from = index,
    along.with = comments_summary$reference_position
  )
  comments_summary$new_id <- as.character(comments_summary$new_id)
  for (i in seq_len(nrow(comments_summary))) {
    content_section <- comments_summary[[i, "content_section"]]
    fake_id <- comments_summary[[i, "fake_id"]]
    new_id <- comments_summary[[i, "new_id"]]
    content_section <- gsub(
      fake_id,
      new_id,
      content_section,
      fixed = TRUE
    )
    newcomment <- as_xml_document(content_section)
    xml_add_child(comments$get(), newcomment)

    reference_position <- comments_summary[[i, "reference_position"]]
    range_start_position <- comments_summary[[i, "range_start_position"]]
    range_end_position <- comments_summary[[i, "range_end_position"]]
    xml_str[c(reference_position, range_start_position, range_end_position)] <-
      gsub(
        fake_id,
        new_id,
        xml_str[c(reference_position, range_start_position, range_end_position)]
      )
  }

  # drop added comments contents from document xml
  lines_index_to_drop <- mapply(
    function(start, stop, str) {
      start:stop
    },
    start = comments_summary$start,
    stop = comments_summary$stop,
    SIMPLIFY = FALSE
  )
  lines_index_to_drop <- unlist(lines_index_to_drop)
  if (length(lines_index_to_drop) > 0) {
    xml_str <- xml_str[-lines_index_to_drop]
  }

  xml_str
}

extract_comments <- function(x, xml_str, index) {
  # comments extraction
  starts <- grep("<w:comment ", xml_str, fixed = TRUE)
  stops <- grep("</w:comment>", xml_str, fixed = TRUE)

  reference_positions <- grep(
    "<w:commentReference w:officer=\"true\"",
    xml_str,
    fixed = FALSE
  )
  range_start_positions <- grep(
    "<w:commentRangeStart( xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"){0,1} w:officer=\"true\"",
    xml_str,
    fixed = FALSE
  )
  range_end_positions <- grep(
    "<w:commentRangeEnd( xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"){0,1} w:officer=\"true\"",
    xml_str,
    fixed = FALSE
  )

  starts <- starts[(starts - 1) %in% reference_positions]
  stops <- stops[(starts - 1) %in% reference_positions]

  content_chrs <- mapply(
    function(start, stop, str) {
      z <- str[start:stop]
      paste(z, collapse = "")
    },
    start = starts,
    stop = stops,
    MoreArgs = list(str = xml_str),
    SIMPLIFY = FALSE
  )
  content_chrs <- unlist(content_chrs) %||% character(0)

  ids <- sprintf(
    "%.0f",
    seq(from = index, along.with = content_chrs)
  )

  if (length(ids) < 1) {
    reference_positions <- integer(0)
    range_start_positions <- integer(0)
    range_end_positions <- integer(0)
  }

  comment_sections <- data.frame(
    content_section = content_chrs,
    id = ids,
    start = starts,
    stop = stops,
    reference_position = reference_positions,
    range_start_position = range_start_positions,
    range_end_position = range_end_positions,
    stringsAsFactors = FALSE
  )
  comment_sections
}
extract_footnotes <- function(x, xml_str, index) {
  # init next_rel_id that will be the first rid
  relationships <- x$relationship()
  next_rel_id <- relationships$get_next_id()

  # footers extraction
  starts <- grep("<w:footnote ", xml_str, fixed = TRUE)
  stops <- grep("</w:footnote>", xml_str, fixed = TRUE)

  reference_positions <- grep(
    "<w:footnoteReference w:officer=\"true\"",
    xml_str,
    fixed = TRUE
  )

  starts <- starts[(starts - 1) %in% reference_positions]
  stops <- stops[(starts - 1) %in% reference_positions]

  content_chrs <- mapply(
    function(start, stop, str) {
      z <- str[start:stop]
      paste(z, collapse = "")
    },
    start = starts,
    stop = stops,
    MoreArgs = list(str = xml_str),
    SIMPLIFY = FALSE
  )
  content_chrs <- unlist(content_chrs) %||% character(0)

  rids <- sprintf(
    "%.0f",
    seq(from = index, along.with = content_chrs)
  )

  if (length(rids) < 1) {
    reference_positions <- integer(0)
  }
  footnote_sections <- data.frame(
    content_section = content_chrs,
    rid = rids,
    start = starts,
    stop = stops,
    reference_position = reference_positions,
    stringsAsFactors = FALSE
  )
  footnote_sections
}
