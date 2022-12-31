#' @export
#' @title add images into an rdocx object
#' @description reference images into a Word document.
#' This function is now useless as the processing of images
#' is automated when using [print.rdocx()].
#'
#' @param x an rdocx object
#' @param src a vector of character containing image filenames.
#' @family functions for officer extensions
#' @keywords internal
docx_reference_img <- function(x, src) {
  x
}

#' @export
#' @title transform an xml string with images references
#' @description This function is useless now, as the processing of images
#' is automated when using [print.rdocx()].
#' @param x an rdocx object
#' @param str wml string
#' @family functions for officer extensions
#' @keywords internal
wml_link_images <- function(x, str) {
  str
}

# by capturing the path, we are making 'unique' new image names.
#' @importFrom openssl sha1
fake_newname <- function(filename) {
  which_files <- grepl("\\.[a-zA-Z0-0]+$", filename)
  file_type <- gsub("(.*)(\\.[a-zA-Z0-0]+)$", "\\2", filename[which_files])
  dest_basename <- sapply(filename[which_files], function(z) {
    as.character(sha1(file(z)))
  }
  )
  dest_basename <- paste0(dest_basename, file_type)
  x <- filename
  x[which_files] <- dest_basename
  x
}

process_images <- function(doc_obj, relationships, package_dir, media_dir = "word/media", media_rel_dir = "media") {
  hl_nodes <- xml_find_all(
    doc_obj$get(), "//a:blip[@r:embed]|//asvg:svgBlip[@r:embed]",
    ns = c(
      "a" = "http://schemas.openxmlformats.org/drawingml/2006/main",
      "asvg" = "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
      "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
  )
  which_to_add <- hl_nodes[!grepl("^rId[0-9]+$", xml_attr(hl_nodes, "embed"))]
  hl_ref <- unique(xml_attr(which_to_add, "embed"))
  for (i in seq_along(hl_ref)) {


    dest_basename <- fake_newname(hl_ref[i])
    img_path <- file.path(package_dir, media_dir)
    if (!file.exists(file.path(img_path, dest_basename))) {
      dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
      file.copy(from = hl_ref[i], to = file.path(img_path, dest_basename))
    }
    if (!file.path(media_rel_dir, dest_basename) %in% relationships$get_data()$target){
      rid <- sprintf("rId%.0f", relationships$get_next_id())
      relationships$add(
        id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target = file.path(media_rel_dir, dest_basename)
      )
    } else {
      reldf <- relationships$get_data()
      rid <- reldf$id[basename(reldf$target) %in% dest_basename]
    }
    which_match_id <- grepl(dest_basename, fake_newname(xml_attr(which_to_add, "embed")), fixed = TRUE)
    xml_attr(which_to_add[which_match_id], "r:embed") <- rep(rid, sum(which_match_id))
  }
}
process_docx_poured <- function(doc_obj, relationships, content_type,
                                package_dir,
                                media_dir = "word", media_rel_dir = "./") {
  hl_nodes <- xml_find_all(
    doc_obj$get(), "//w:altChunk[@r:id]",
    ns = c(
      "w" = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
  )

  which_to_add <- hl_nodes[!grepl("^rId[0-9]+$", xml_attr(hl_nodes, "id"))]
  hl_ref <- unique(xml_attr(which_to_add, "id"))
  for (i in seq_along(hl_ref)) {
    rid <- sprintf("rId%.0f", relationships$get_next_id())

    img_path <- file.path(package_dir, media_dir)
    file.copy(from = hl_ref[i], to = file.path(package_dir, media_dir, basename(hl_ref[i])))

    relationships$add(
      id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
      target = file.path(media_rel_dir, basename(hl_ref[i]))
    )
    content_type$add_override(
      setNames("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", paste0("/", media_dir, "/", basename(hl_ref[i])))
    )

    which_match_id <- grepl(hl_ref[i], xml_attr(which_to_add, "id"), fixed = TRUE)
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(rid, sum(which_match_id))
  }
}

#' @importFrom xml2 xml_remove as_xml_document xml_parent xml_child
process_footnotes <- function(x) {
  footnotes <- x$footnotes
  doc_obj <- x$doc_obj

  rel <- doc_obj$relationship()

  hl_nodes <- xml_find_all(doc_obj$get(), "//w:footnoteReference[@w:id]")
  which_to_add <- hl_nodes[grepl("^footnote", xml_attr(hl_nodes, "id"))]
  hl_ref <- xml_attr(which_to_add, "id")
  for (i in seq_along(hl_ref)) {
    next_id <- rel$get_next_id()
    rel$add(
      paste0("rId", next_id),
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
      target = "footnotes.xml"
    )

    index <- length(xml_find_all(footnotes$get(), "w:footnote")) - 1
    xml_attr(which_to_add[[i]], "w:id") <- index

    run <- xml_parent(which_to_add[[i]])

    run_rstyle <- xml_child(run, "w:rPr/w:rStyle")

    styles <- styles_info(x, type = "character")
    style_id <- xml_attr(run_rstyle, "val")
    style_id <- styles$style_id[styles$style_name %in% style_id]

    xml_attr(run_rstyle, "w:val") <- style_id

    footnote <- xml_child(which_to_add[[i]], "w:footnote")
    xml_attr(footnote, "w:id") <- index

    footnote_rstyle <- xml_child(footnote, "w:p/w:r/w:rPr/w:rStyle")
    xml_attr(footnote_rstyle, "w:val") <- style_id

    newfootnote <- as_xml_document(as.character(footnote))
    xml_remove(footnote)

    xml_add_child(footnotes$get(), newfootnote)
  }
}
process_links <- function(doc_obj, type = "wml") {
  rel <- doc_obj$relationship()
  if ("wml" %in% type) {
    hl_nodes <- xml_find_all(doc_obj$get(), "//w:hyperlink[@r:id]")
  } else {
    hl_nodes <- xml_find_all(doc_obj$get(), "//a:hlinkClick[@r:id]")
  }
  which_to_add <- hl_nodes[!grepl("^rId[0-9]+$", xml_attr(hl_nodes, "id"))]
  hl_ref <- unique(xml_attr(which_to_add, "id"))
  for (i in seq_along(hl_ref)) {
    rid <- sprintf("rId%.0f", rel$get_next_id())

    rel$add(
      id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = officer_url_decode(hl_ref[i]), target_mode = "External"
    )

    which_match_id <- grepl(hl_ref[i], xml_attr(which_to_add, "id"), fixed = TRUE)
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(rid, sum(which_match_id))
  }
}
