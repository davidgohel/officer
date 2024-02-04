# Adds notes master if not present. There can be at most one notesMaster in a presentation.
add_notesMaster <- function(x) {

  presentation_relations <- x$presentation$rel_df()
  if (!any(grepl("notesMaster$", presentation_relations$type))) {

    # copy notesMaster file from templates
    path <- system.file(package = "officer", "template/notesMaster.xml")
    xml_file <- file.path(x$package_dir, "ppt/notesMasters/notesMaster1.xml")
    dir.create(dirname(xml_file), showWarnings = FALSE, recursive = FALSE)
    file.copy(path, to = xml_file,
              overwrite = TRUE,
              copy.mode = FALSE)

    # copy theme file
    theme_file <- "theme1.xml"
    theme_files <- list.files(path = file.path(x$package_dir, "ppt/theme"), pattern = "\\.xml$", full.names = F)
    theme_index <- as.integer(gsub("^(theme)([0-9]+)(\\.xml)$", "\\2", theme_files ))
    theme_file <- gsub(pattern = "[0-9]+", replacement = max(theme_index) + 1, theme_file)
    file.copy(file.path(x$package_dir, "ppt/theme/theme1.xml"),
              to = file.path(x$package_dir, "ppt/theme", theme_file),
              overwrite = TRUE,
              copy.mode = FALSE)

    # add relation to theme file
    newrel <- relationship$new()$add(id = "rId1",
                                     type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                                     target = paste0("../theme/", theme_file))
    newrel$write(path = file.path(dirname(xml_file), "_rels",  paste0(basename(xml_file), ".rels")))

    x$notesMaster$add_notesMaster(xml_file)

    # register xml files in content_type
    partname <- "/ppt/notesMasters/notesMaster1.xml"
    content_type <- setNames("application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml", partname )
    x$content_type$add_override(content_type)

    partname <- paste0("/ppt/theme/", theme_file)
    content_type <- setNames("application/vnd.openxmlformats-officedocument.theme+xml", partname )
    x$content_type$add_override(content_type)

    # add notesMaster relation to presentation
    prs_relations <- x$presentation$relationship()
    rId <-  paste0("rId",prs_relations$get_next_id())
    prs_relations$add(id = rId,
                      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
                      target = "notesMasters/notesMaster1.xml")
    prs_relations$write(file.path(x$package_dir, "ppt/_rels", "presentation.xml.rels"))

    xml_elt <- as_xml_document(paste0("<p:notesMasterIdLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">
                                         <p:notesMasterId r:id=\"", rId, "\"/>
                                       </p:notesMasterIdLst>"))
    xml_add_sibling(xml_find_first(x$presentation$get(), "//p:sldMasterIdLst"), xml_elt)
  }
  return(x)
}

# Create notesSlide for current slide if it doesn't exist.
# Returns name of notesSlide for current slide.
get_or_create_notesSlide <- function(x) {

  slide <- x$slide$get_slide(x$cursor)
  rels <- slide$rel_df()
  if (!any(grepl("notesSlide$", rels$type))) {

    # add notesMaster if it doesn't exist
    add_notesMaster(x)

    nslidename <- x$notesSlide$get_new_slidename()

    # copy slide template
    xml_file <- file.path(x$package_dir, "ppt/notesSlides", nslidename)
    path <- system.file(package = "officer", "template/notesSlide.xml")
    file.copy(path, to = xml_file, copy.mode = FALSE)

    # add relationships to notesMaster and slide
    rel_filename <- file.path(
      dirname(xml_file), "_rels",
      paste0(basename(xml_file), ".rels"))
    newrel <- relationship$new()$add(
      id = "rId1",
      type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
      target = "../notesMasters/notesMaster1.xml")
    newrel$add(id = "rId2",
               type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
               target = file.path("../slides", slide$name()))
    newrel$write(path = rel_filename)

    # add relationship from slide to notesSlide
    rel <- slide$relationship()
    rel$add(id = paste0("rId", rel$get_next_id()),
            type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
            target = file.path( "../notesSlides", nslidename) )
    rel$write(file.path(
      dirname(slide$file_name()), "_rels",
      paste0(basename(slide$file_name()), ".rels")))

    # update presentation elements
    x$content_type$add_notesSlide(partname = file.path( "/ppt/notesSlides", nslidename) )

    # add new notesSlide to collection
    x$notesSlide$add_slide(xml_file)
  } else {
    nslidename <- basename(rels$target[grepl("notesSlide$", rels$type)])
  }
  return(nslidename)
}



#' @export
#' @title Location of a named placeholder for notes
#' @description The function will use the label of a placeholder
#' to find the corresponding location in the slide notes.
#' @param ph_label placeholder label of the used notes master
#' @param ... unused arguments
notes_location_label <- function( ph_label, ...){
  x <- list(ph_label = ph_label)
  class(x) <- c("location_label", "location_str")
  x
}

#' @export
#' @title Location of a placeholder for notes
#' @description The function will use the type name of the placeholder (e.g. body, hdr),
#' to find the corresponding location.
#' @param type placeholder label of the used notes master
#' @param ... unused arguments
notes_location_type <- function( type = "body", ...){

  ph_types <- c("ftr", "sldNum", "hdr", "body", "sldImg")
  if(!type %in% ph_types){
    stop("argument type ('", type, "') expected to be a value of ",
         paste0(shQuote(ph_types), collapse = ", "), ".")
  }
  x <- list(type = type)
  class(x) <- c("location_type", "location_str")
  x
}

ph_from_location <- function(loc, doc, ...){
  UseMethod("ph_from_location", loc)
}


ph_from_location.location_label <- function(loc, doc, ...) {
  xfrm <- doc$notesMaster$xfrm()
  location <- xfrm[xfrm$ph_label == loc$ph_label, ]
  if (nrow(location) < 1) stop("No placeholder with label ", loc$ph_label, " found!")

  id <- uuid_generate()
  str <- "<p:nvSpPr><p:cNvPr id=\"%s\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/>"
  new_ph <- sprintf(str, id, location$ph_label, location$ph)
  return(list(ph = new_ph, label = location$ph_label))
}


ph_from_location.location_type <- function(loc, doc, ...) {
  xfrm <- doc$notesMaster$xfrm()
  location <- xfrm[xfrm$type == loc$type, ]
  if (nrow(location) < 1) stop("No placeholder of type ", loc$type, " found!")

  id <- uuid_generate()
  str <- "<p:nvSpPr><p:cNvPr id=\"%s\" name=\"%s\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr/>"
  new_ph <- sprintf(str, id, location$ph_label[1], location$ph[1])
  return(list(ph = new_ph, label = location$ph_label[1]))
}

#' @export
#' @title Set notes for current slide
#' @description Set speaker notes for the current slide in a pptx presentation.
#' @param x an rpptx object
#' @param value text to be added to notes
#' @param location a placeholder location object.
#' It will be used to specify the location of the new shape. This location
#' can be defined with a call to one of the notes_ph functions. See
#' section \code{"see also"}.
#' @param ... further arguments passed to or from other methods.
#' @examples
#' # this name will be used to print the file
#' # change it to "youfile.pptx" to write the pptx
#' # file in your working directory.
#' fileout <- tempfile(fileext = ".pptx")
#' fpt_blue_bold <- fp_text_lite(color = "#006699", bold = TRUE)
#' doc <- read_pptx()
#' # add a slide with some text ----
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with(x = doc, value = "Slide Title 1",
#'    location = ph_location_type(type = "title") )
#' # set speaker notes for the slide ----
#' doc <- set_notes(doc, value = "This text will only be visible for the speaker.",
#'    location = notes_location_type("body"))
#'
#' # add a slide with some text ----
#' doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
#' doc <- ph_with(x = doc, value = "Slide Title 2",
#'    location = ph_location_type(type = "title") )
#' bl <- block_list(
#'   fpar(ftext("hello world", fpt_blue_bold)),
#'   fpar(ftext("Turlututu chapeau pointu", fpt_blue_bold))
#' )
#' doc <- set_notes(doc, value = bl,
#'    location = notes_location_type("body"))
#'
#' print(doc, target = fileout)
#'
#' @seealso [print.rpptx()], [read_pptx()], [add_slide()], [notes_location_label()], [notes_location_type()]
#' @family functions slide manipulation
set_notes <- function(x, value, location, ...){
  UseMethod("set_notes", value)
}

#' @export
#' @describeIn set_notes add a character vector to a place holder in the notes on the
#' current slide, values will be added as paragraphs.
set_notes.character <- function( x, value, location, ... ){

  # get or create notesSlide for current slide
  nslidename <- get_or_create_notesSlide(x)

  idx <- x$notesSlide$slide_index(nslidename)
  nSlide <- x$notesSlide$get_slide(idx)

  new_ph <- ph_from_location(location, x)

  # remove placeholder if already used
  xml_remove(xml_find_first(nSlide$get(), paste0("//p:spTree/p:sp[p:nvSpPr/p:cNvPr[@name='", new_ph$label, "']]")))


  pars <- paste0("<a:p><a:r><a:rPr/><a:t>", htmlEscapeCopy(value), "</a:t></a:r></a:p>", collapse = "")
  xml_elt <- as_xml_document(paste0( psp_ns_yes, new_ph$ph,
                                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                                     pars, "</p:txBody></p:sp>" ))

  xml_add_child(xml_find_first(nSlide$get(), "//p:spTree"), xml_elt)
  nSlide$fortify_id()

  return(x)
}

#' @export
#' @describeIn set_notes add a [block_list()] to a place holder in the notes on the
#' current slide.
set_notes.block_list <- function( x, value, location, ... ){

  # get or create notesSlide for current slide
  nslidename <- get_or_create_notesSlide(x)

  idx <- x$notesSlide$slide_index(nslidename)
  nSlide <- x$notesSlide$get_slide(idx)

  new_ph <- ph_from_location(location, x)

  # remove placeholder if already used
  xml_remove(xml_find_first(nSlide$get(), paste0("//p:spTree/p:sp[p:nvSpPr/p:cNvPr[@name='", new_ph$label, "']]")))

  pars <- sapply(value, to_pml)
  pars <- paste0(pars, collapse = "")

  xml_elt <- as_xml_document(paste0( psp_ns_yes, new_ph$ph,
                                     "<p:txBody><a:bodyPr/><a:lstStyle/>",
                                     pars, "</p:txBody></p:sp>" ))

  xml_add_child(xml_find_first(nSlide$get(), "//p:spTree"), xml_elt)
  nSlide$fortify_id()

  return(x)
}
