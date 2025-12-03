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
  })
  dest_basename <- paste0(dest_basename, file_type)
  x <- filename
  x[which_files] <- dest_basename
  x
}

process_images <- function(
  doc_obj,
  relationships,
  package_dir,
  media_dir = "word/media",
  media_rel_dir = "media"
) {
  hl_nodes <- xml_find_all(
    doc_obj$get(),
    "//a:blip[@r:embed]|//asvg:svgBlip[@r:embed]",
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
    which_match_id <- grepl(
      dest_basename,
      fake_newname(xml_attr(which_to_add, "embed")),
      fixed = TRUE
    )
    xml_attr(which_to_add[which_match_id], "r:embed") <- rep(
      rid,
      sum(which_match_id)
    )
  }
}

process_docx_poured <- function(
  doc_obj,
  relationships,
  content_type,
  package_dir,
  media_dir = "word"
) {
  hl_nodes <- xml_find_all(
    doc_obj$get(),
    "//w:altChunk[@r:id]",
    ns = c(
      "w" = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
  )

  which_to_add <- hl_nodes[!grepl("^rId[0-9]+$", xml_attr(hl_nodes, "id"))]
  hl_ref <- unique(xml_attr(which_to_add, "id"))
  for (i in seq_along(hl_ref)) {
    rid <- sprintf("rId%.0f", relationships$get_next_id())

    file.copy(
      from = hl_ref[i],
      to = file.path(package_dir, media_dir, basename(hl_ref[i]))
    )

    relationships$add(
      id = rid,
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
      target = basename(hl_ref[i])
    )
    content_type$add_override(
      setNames(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        paste0("/", media_dir, "/", basename(hl_ref[i]))
      )
    )

    which_match_id <- grepl(
      hl_ref[i],
      xml_attr(which_to_add, "id"),
      fixed = TRUE
    )
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(
      rid,
      sum(which_match_id)
    )
  }
}

#' @importFrom xml2 xml_remove as_xml_document xml_parent xml_child
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
      id = rid,
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = officer_url_decode(hl_ref[i]),
      target_mode = "External"
    )

    which_match_id <- grepl(
      hl_ref[i],
      xml_attr(which_to_add, "id"),
      fixed = TRUE
    )
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(
      rid,
      sum(which_match_id)
    )
  }
}


update_hf_list <- function(part_list = list(), type = "header", package_dir) {
  files <- list.files(
    path = file.path(package_dir, "word"),
    pattern = sprintf("^%s[0-9]*.xml$", type)
  )
  files <- files[!basename(files) %in% names(part_list)]
  if (type %in% "header") {
    cursor <- "/w:hdr/*[1]"
    body_xpath <- "/w:hdr"
  } else {
    cursor <- "/w:ftr/*[1]"
    body_xpath <- "/w:ftr"
  }

  new_list <- lapply(files, function(x) {
    docx_part$new(
      path = package_dir,
      main_file = x,
      cursor = cursor,
      body_xpath = body_xpath
    )
  })
  names(new_list) <- basename(files)
  append(part_list, new_list)
}

#' @export
#' @title Remove unused media from a document
#' @description The function will scan the media
#' directory and delete images that are not used
#' anymore. This function is to be used when images
#' have been replaced many times.
#' @param x `rdocx` or `rpptx` object
#' @param warn_user TRUE to make sure users are warned when
#' using this function that will be un-exported in a next
#' version
#' @keywords internal
sanitize_images <- function(x, warn_user = TRUE) {

  if (warn_user) {
    cli::cli_warn(
      c(
        "!" = "Function {.fn sanitize_images} will be removed soon. You should remove calls to {.fn sanitize_images}.",
        "i" = "Instead the function is run before saving the document to a file."
      )
    )
  }

  if (inherits(x, "rdocx")) {

    image_files <- c()
    all_docs <- append(x$headers, x$footers)
    all_docs[[length(all_docs)+1]] <- x$doc_obj
    all_docs[[length(all_docs)+1]] <- x$footnotes

    for (doc_part in all_docs) {
      suppressWarnings({
        blip_nodes <- xml_find_all(
          doc_part$get(),
          "//a:blip[contains(@r:embed, 'rId')]|//asvg:svgBlip[contains(@r:embed, 'rId')]",
          ns = c(
            "a" = "http://schemas.openxmlformats.org/drawingml/2006/main",
            "asvg" = "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
            "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          )
        )
      })

      embed_list <- xml_attr(blip_nodes, "embed")
      embed_data <- filter(
        .data = doc_part$rel_df(),
        basename(.data$type) %in% "image",
        .data$id %in% embed_list
      )
      embed_data <- embed_data$target
      image_files[[length(image_files)+1]] <- embed_data
    }

    image_files <- do.call(c, image_files)
    image_files <- unique(image_files)


    base_doc <- file.path(x$package_dir, "word")
    existing_img <- list.files(
      file.path(base_doc, "media"),
      pattern = "\\.(png|jpg|jpeg|eps|emf|svg)$",
      ignore.case = TRUE,
      recursive = TRUE,
      full.names = TRUE
    )

    existing_img <- gsub(paste0(base_doc, "/"), "", existing_img, fixed = TRUE)
    unlink(
      file.path(base_doc, setdiff(existing_img, image_files)),
      force = TRUE
    )
  } else if (inherits(x, "rpptx")) {

    rel_files <- list.files(
      x$package_dir,
      pattern = "\\.xml.rels$",
      recursive = TRUE,
      full.names = TRUE
    )

    image_files <- lapply(rel_files, function(x) {
      zz <- read_xml(x)
      rels <- xml_children(zz)
      rels <- rels[
        xml_attr(rels, "Type") %in%
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      ]
      xml_attr(rels, "Target")
    })
    image_files <- unique(unlist(image_files))
    base_doc <- file.path(x$package_dir, "ppt")
    existing_img <- list.files(
      file.path(base_doc, "media"),
      pattern = "\\.(png|jpg|jpeg|eps|emf|svg)$",
      ignore.case = TRUE,
      recursive = TRUE,
      full.names = TRUE
    )
    existing_img <- gsub(paste0(base_doc, "/"), "../", existing_img, fixed = TRUE)
    unlink(
      file.path(base_doc, setdiff(existing_img, image_files)),
      force = TRUE
    )
  }
  x
}
