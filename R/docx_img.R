#' @export
#' @title add images into an rdocx object
#' @description reference images into a Word document.
#' This function is to be used with \code{\link{wml_link_images}}.
#'
#' Images need to be referenced into the
#' Word document, this will generate unique
#' identifiers that need to be known to
#' link these images with their corresponding xml code (wml).
#'
#' @param x an rdocx object
#' @param src a vector of character containing image filenames.
#' @family functions for officer extensions
#' @keywords internal
docx_reference_img <- function(x, src) {
  # x <- part_reference_img(x, src, "doc_obj")
  x
}

#' @export
#' @title transform an xml string with images references
#' @description The function replace images filenames
#' in an xml string with their id. The wml code cannot
#' be valid without this operation.
#' @details
#' The function is available to allow the creation of valid
#' wml code containing references to images.
#' @param x an rdocx object
#' @param str wml string
#' @family functions for officer extensions
#' @keywords internal
wml_link_images <- function(x, str) {
  str
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
    rid <- sprintf("rId%.0f", relationships$get_next_id())

    img_path <- file.path(package_dir, media_dir)
    dir.create(img_path, recursive = TRUE, showWarnings = FALSE)
    file.copy(from = hl_ref[i], to = file.path(package_dir, media_dir, basename(hl_ref[i])))

    relationships$add(
      id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target = file.path(media_rel_dir, basename(hl_ref[i]))
    )

    which_match_id <- grepl(hl_ref[i], xml_attr(which_to_add, "embed"), fixed = TRUE)
    xml_attr(which_to_add[which_match_id], "r:embed") <- rep(rid, sum(which_match_id))
  }
}
