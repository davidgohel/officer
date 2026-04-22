# worksheets ------------------------------------------------------------

worksheets <- R6Class(
  "worksheets",
  inherit = openxml_document,

  public = list(
    initialize = function(path) {
      super$initialize(character(0))
      private$package_dir <- path
      presentation_filename <- file.path(path, "xl", "workbook.xml")
      self$feed(presentation_filename)

      slide_df <- self$get_sheets_df()
      private$sheet_id <- slide_df$sheet_id
      private$sheet_rid <- slide_df$rid
      private$sheet_name <- slide_df$name
    },

    view_on_sheet = function(sheet) {
      sheet_id <- self$get_sheet_id(sheet)
      wb_view <- xml_find_first(self$get(), "d1:bookViews/d1:workbookView")
      xml_attr(wb_view, "activeTab") <- sheet_id - 1
      self$save()
    },

    add_sheet = function(target, label) {
      private$rels_doc$add(
        id = paste0("rId", private$rels_doc$get_next_id()),
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        target = target
      )
      rels <- private$rels_doc$get_data()
      rid <- rels[rels$target %in% target, "id"]

      ids <- private$sheet_id
      new_id <- max(ids) + 1

      private$sheet_id <- c(private$sheet_id, new_id)
      private$sheet_rid <- c(private$sheet_rid, rid)
      private$sheet_name <- c(private$sheet_name, label)

      children_ <- xml_children(self$get())
      sheets_id <- which(sapply(children_, function(x) xml_name(x) == "sheets"))
      xml_list <- xml_children(children_[[sheets_id]])

      xml_elt <- paste(
        sprintf(
          "<sheet name=\"%s\" sheetId=\"%.0f\" r:id=\"%s\"/>",
          htmlEscapeCopy(private$sheet_name),
          private$sheet_id,
          private$sheet_rid
        ),
        collapse = ""
      )
      xml_elt <- paste0(
        "<sheets xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
        xml_elt,
        "</sheets>"
      )
      xml_elt <- as_xml_document(xml_elt)

      if (!inherits(xml_list, "xml_missing")) {
        xml_replace(children_[[sheets_id]], xml_elt)
      } else {
        cli::cli_abort("Could not find sheets entity in workbook XML.")
      }

      self
    },

    get_new_sheetname = function() {
      sheet_dir <- file.path(private$package_dir, "xl/worksheets")
      if (!file.exists(sheet_dir)) {
        dir.create(
          file.path(sheet_dir, "_rels"),
          showWarnings = FALSE,
          recursive = TRUE
        )
      }

      sheet_files <- basename(list.files(sheet_dir, pattern = "\\.xml$"))
      sheet_name <- "sheet1.xml"
      if (length(sheet_files)) {
        slide_index <- as.integer(gsub(
          "^(sheet)([0-9]+)(\\.xml)$",
          "\\2",
          sheet_files
        ))
        sheet_name <- gsub(
          pattern = "[0-9]+",
          replacement = max(slide_index) + 1,
          sheet_name
        )
      }
      sheet_name
    },

    sheet_names = function() {
      private$sheet_name
    },

    get_sheet_id = function(name) {
      sheets_df <- self$get_sheets_df()
      bool_name_in_list <- sheets_df$name %in% name
      n_matches <- sum(bool_name_in_list, na.rm = TRUE)
      if (n_matches < 1) {
        cli::cli_abort("Could not find sheet {.val {name}}.")
      }
      sheets_df$sheet_id[bool_name_in_list]
    },

    get_sheets_df = function() {
      children_ <- xml_children(self$get())
      sheets_id <- which(sapply(children_, function(x) xml_name(x) == "sheets"))
      sheet_nodes <- xml_children(children_[[sheets_id]])
      data.frame(
        stringsAsFactors = FALSE,
        name = xml_attr(sheet_nodes, "name"),
        sheet_id = as.integer(xml_attr(sheet_nodes, "sheetId")),
        rid = xml_attr(sheet_nodes, "id")
      )
    }
  ),
  private = list(
    sheet_id = NULL,
    sheet_rid = NULL,
    sheet_name = NULL,
    package_dir = NULL
  )
)

# sheet ------------------------------------------------------------
sheet <- R6Class(
  "sheet",
  inherit = openxml_document,
  public = list(
    feed = function(file) {
      private$filename <- file
      private$doc <- read_xml(file)

      private$rels_filename <- file.path(
        dirname(file),
        "_rels",
        paste0(basename(file), ".rels")
      )

      if (file.exists(private$rels_filename)) {
        private$rels_doc <- relationship$new()$feed_from_xml(
          private$rels_filename
        )
      } else {
        new_rel <- relationship$new()
        new_rel$write(private$rels_filename)
        private$rels_doc <- new_rel
      }
      self
    }
  )
)


# xlsx_drawing ----
#' @export
#' @title Drawing manager for xlsx sheets
#' @description R6 class that manages a drawing file for an xlsx
#' sheet. Used internally by [sheet_add_drawing()] methods.
#' @param package_dir path to the unpacked xlsx directory
#' @param sheet_obj a sheet R6 object
#' @param content_type a content_type R6 object
#' @param chart_rid relationship id of the chart
#' @param from_col top-left column anchor (0-based)
#' @param from_row top-left row anchor (0-based)
#' @param to_col bottom-right column anchor (0-based)
#' @param to_row bottom-right row anchor (0-based)
#' @param chart_basename filename of the chart XML
#' @keywords internal
xlsx_drawing <- R6Class(
  "xlsx_drawing",
  inherit = openxml_document,

  public = list(
    #' @description Create or reuse a drawing for a sheet.
    initialize = function(package_dir, sheet_obj, content_type) {
      super$initialize("xl/drawings")
      private$package_dir <- package_dir

      drawings_dir <- file.path(package_dir, "xl", "drawings")
      dir.create(drawings_dir, recursive = TRUE, showWarnings = FALSE)
      rels_dir <- file.path(drawings_dir, "_rels")
      dir.create(rels_dir, recursive = TRUE, showWarnings = FALSE)

      # check if sheet already has a drawing
      sheet_rels <- sheet_obj$rel_df()
      drawing_rel <- sheet_rels[
        grepl("drawing$", sheet_rels$type),
        ,
        drop = FALSE
      ]

      if (nrow(drawing_rel) > 0) {
        # reuse existing drawing
        drawing_target <- drawing_rel$target[1L]
        drawing_file <- file.path(
          package_dir,
          "xl",
          "worksheets",
          drawing_target
        )
        self$feed(normalizePath(drawing_file, mustWork = TRUE))
      } else {
        # create new drawing
        drawing_name <- private$next_drawing_name(drawings_dir)
        drawing_file <- file.path(drawings_dir, drawing_name)

        drawing_xml <- paste0(
          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
          "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"",
          " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
        )
        writeLines(drawing_xml, drawing_file, useBytes = TRUE)
        self$feed(drawing_file)

        # add relationship from sheet to drawing
        next_id <- sheet_obj$relationship()$get_next_id()
        rid <- paste0("rId", next_id)
        sheet_obj$relationship()$add(
          id = rid,
          type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          target = paste0("../drawings/", drawing_name)
        )

        # add <drawing r:id="rIdN"/> to sheet XML
        ns <- c(
          d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        )
        existing_drawing <- xml_find_first(
          sheet_obj$get(),
          "d1:drawing",
          ns = ns
        )
        if (inherits(existing_drawing, "xml_missing")) {
          drawing_ref <- read_xml(sprintf(
            "<drawing xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"%s\"/>",
            rid
          ))
          # insert before extLst if present (OOXML element order)
          ext_lst <- xml_find_first(
            sheet_obj$get(),
            "d1:extLst",
            ns = ns
          )
          if (!inherits(ext_lst, "xml_missing")) {
            xml_add_sibling(ext_lst, drawing_ref, .where = "before")
          } else {
            xml_add_child(sheet_obj$get(), drawing_ref)
          }
        }
        sheet_obj$save()

        # content type
        partname <- paste0("/xl/drawings/", drawing_name)
        content_type$add_override(setNames(
          "application/vnd.openxmlformats-officedocument.drawing+xml",
          partname
        ))
      }
    },

    #' @description Add a chart anchor to the drawing.
    add_chart_anchor = function(
      chart_rid,
      from_col = 3L,
      from_row = 1L,
      to_col = 10L,
      to_row = 15L
    ) {
      anchor_xml <- sprintf(
        paste0(
          "<xdr:twoCellAnchor",
          " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"",
          " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
          " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
          " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">",
          "<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>",
          "<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
          "<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>",
          "<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
          "<xdr:graphicFrame macro=\"\">",
          "<xdr:nvGraphicFramePr>",
          "<xdr:cNvPr id=\"%d\" name=\"Chart %d\"/>",
          "<xdr:cNvGraphicFramePr/>",
          "</xdr:nvGraphicFramePr>",
          "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>",
          "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">",
          "<c:chart r:id=\"%s\"/>",
          "</a:graphicData></a:graphic>",
          "</xdr:graphicFrame>",
          "<xdr:clientData/>",
          "</xdr:twoCellAnchor>"
        ),
        from_col,
        from_row,
        to_col,
        to_row,
        private$next_cNvPr_id(),
        private$next_cNvPr_id(),
        chart_rid
      )

      xml_add_child(self$get(), read_xml(anchor_xml))
      self$save()
      self
    },

    #' @description Add a relationship from the drawing to a chart file.
    add_chart_rel = function(chart_basename) {
      next_id <- self$relationship()$get_next_id()
      rid <- paste0("rId", next_id)
      self$relationship()$add(
        id = rid,
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        target = paste0("../charts/", chart_basename)
      )
      self$save()
      rid
    },

    #' @description Add a relationship from the drawing to a media image.
    #' @param image_basename filename (without directory) of the image
    #'   sitting in `xl/media/`.
    add_image_rel = function(image_basename) {
      next_id <- self$relationship()$get_next_id()
      rid <- paste0("rId", next_id)
      self$relationship()$add(
        id = rid,
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target = paste0("../media/", image_basename)
      )
      self$save()
      rid
    },

    #' @description Add an image anchor to the drawing (absolute placement
    #'   in inches from the top-left corner of the sheet).
    #' @param image_rid relationship id of the image
    #' @param left,top top-left anchor in inches
    #' @param width,height size in inches
    #' @param alt alternative text
    add_image_anchor = function(image_rid,
                                left = 1, top = 1,
                                width = 2, height = 2,
                                alt = "") {
      emu_per_in <- 914400
      nv_id <- private$next_cNvPr_id()
      anchor_xml <- sprintf(
        paste0(
          "<xdr:absoluteAnchor",
          " xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"",
          " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
          " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
          "<xdr:pos x=\"%.0f\" y=\"%.0f\"/>",
          "<xdr:ext cx=\"%.0f\" cy=\"%.0f\"/>",
          "<xdr:pic>",
          "<xdr:nvPicPr>",
          "<xdr:cNvPr id=\"%d\" name=\"Picture %d\" descr=\"%s\"/>",
          "<xdr:cNvPicPr/>",
          "</xdr:nvPicPr>",
          "<xdr:blipFill>",
          "<a:blip r:embed=\"%s\"/>",
          "<a:stretch><a:fillRect/></a:stretch>",
          "</xdr:blipFill>",
          "<xdr:spPr>",
          "<a:xfrm>",
          "<a:off x=\"%.0f\" y=\"%.0f\"/>",
          "<a:ext cx=\"%.0f\" cy=\"%.0f\"/>",
          "</a:xfrm>",
          "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
          "</xdr:spPr>",
          "</xdr:pic>",
          "<xdr:clientData/>",
          "</xdr:absoluteAnchor>"
        ),
        left * emu_per_in, top * emu_per_in,
        width * emu_per_in, height * emu_per_in,
        nv_id, nv_id, htmlEscapeCopy(alt),
        image_rid,
        left * emu_per_in, top * emu_per_in,
        width * emu_per_in, height * emu_per_in
      )
      xml_add_child(self$get(), read_xml(anchor_xml))
      self$save()
      self
    }
  ),

  private = list(
    package_dir = NULL,

    next_drawing_name = function(drawings_dir) {
      existing <- list.files(drawings_dir, pattern = "^drawing[0-9]+\\.xml$")
      if (length(existing) == 0L) {
        return("drawing1.xml")
      }
      nums <- as.integer(gsub("\\D", "", existing))
      paste0("drawing", max(nums) + 1L, ".xml")
    },

    next_cNvPr_id = function() {
      nodes <- xml_find_all(
        self$get(),
        "//xdr:cNvPr",
        ns = c(
          xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        )
      )
      if (length(nodes) == 0L) {
        return(2L)
      }
      max(as.integer(xml_attr(nodes, "id"))) + 1L
    }
  )
)


# dir_sheets ----
dir_sheet <- R6Class(
  "dir_sheet",
  public = list(
    initialize = function(x) {
      private$package_dir <- x$package_dir
      private$sheets_path <- "xl/worksheets"
      self$update()
    },

    update = function() {
      dir_ <- file.path(private$package_dir, private$sheets_path)
      filenames <- list.files(
        path = dir_,
        pattern = "\\.xml$",
        full.names = TRUE
      )

      private$collection <- lapply(
        filenames,
        function(x, path) {
          sheet$new(path)$feed(x)
        },
        private$sheets_path
      )

      names(private$collection) <- sapply(private$collection, function(x) {
        x$name()
      })
      self
    },

    get_sheet_list = function() {
      dir_ <- file.path(private$package_dir, private$sheets_path)
      filenames <- list.files(
        path = dir_,
        pattern = "\\.xml$",
        full.names = TRUE
      )
      sheet_index <- seq_along(filenames)
      if (length(filenames)) {
        filenames <- basename(filenames)
        sheet_index <- as.integer(gsub(
          "^(sheet)([0-9]+)(\\.xml)$",
          "\\2",
          filenames
        ))
        filenames <- filenames[order(sheet_index)]
      }
      filenames
    },

    get_sheet = function(id) {
      l_ <- self$length()
      if (is.null(id) || !between(id, 1, l_)) {
        cli::cli_abort(
          "Invalid sheet id {.val {id}} ({l_} sheet(s) available)."
        )
      }
      sheet_files <- self$get_sheet_list()
      index <- which(names(private$collection) == sheet_files[id])
      private$collection[[index]]
    },

    length = function() {
      length(private$collection)
    }
  ),
  private = list(
    collection = NULL,
    package_dir = NULL,
    sheets_path = NULL
  )
)


# xlsx_styles ----
#' @export
#' @title Style manager for xlsx workbooks
#' @description R6 class that manages fonts, fills, borders, and
#' cell formats in an xlsx workbook's `styles.xml`. Used internally
#' by [sheet_write_data()] and available for extensions.
#' @param package_dir path to the unpacked xlsx directory
#' @param name font family name
#' @param size font size in points
#' @param bold logical, bold
#' @param italic logical, italic
#' @param underline logical, underline
#' @param color hex color string (6 chars, without `#`)
#' @param bg_color hex color string for fill background
#' @param top_style border style for top side
#' @param top_color border color for top side
#' @param bottom_style border style for bottom side
#' @param bottom_color border color for bottom side
#' @param left_style border style for left side
#' @param left_color border color for left side
#' @param right_style border style for right side
#' @param right_color border color for right side
#' @param font_id integer, font index (0-based)
#' @param fill_id integer, fill index (0-based)
#' @param border_id integer, border index (0-based)
#' @param num_fmt_id integer, number format id
#' @param halign horizontal alignment
#' @param valign vertical alignment
#' @param text_rotation text rotation angle (0-180)
#' @param wrap_text logical, enable text wrapping
#' @keywords internal
xlsx_styles <- R6Class(
  "xlsx_styles",

  public = list(
    #' @description Initialize styles from an xlsx package directory.
    initialize = function(package_dir) {
      private$file <- file.path(package_dir, "xl", "styles.xml")
      private$doc <- read_xml(private$file)
      private$ns <- c(
        d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      )
      private$index_existing()
    },

    #' @description Get or create a font index.
    get_font_id = function(
      name = "Calibri",
      size = 11,
      bold = FALSE,
      italic = FALSE,
      underline = FALSE,
      color = "000000"
    ) {
      sig <- paste(
        name,
        size,
        as.integer(bold),
        as.integer(italic),
        as.integer(underline),
        color,
        sep = "|"
      )
      idx <- match(sig, private$font_sigs)
      if (!is.na(idx)) {
        return(idx - 1L)
      }

      xml_parts <- c(
        if (bold) "<b/>",
        if (italic) "<i/>",
        if (underline) "<u/>",
        sprintf("<sz val=\"%s\"/>", size),
        sprintf("<color rgb=\"FF%s\"/>", toupper(color)),
        sprintf("<name val=\"%s\"/>", htmlEscapeCopy(name))
      )
      font_xml <- sprintf(
        "<font xmlns=\"%s\">%s</font>",
        private$ssml,
        paste0(xml_parts, collapse = "")
      )
      private$add_to_collection("d1:fonts", font_xml, sig, "font_sigs")
    },

    #' @description Get or create a fill index.
    get_fill_id = function(bg_color = "FFFFFF") {
      sig <- toupper(bg_color)
      idx <- match(sig, private$fill_sigs)
      if (!is.na(idx)) {
        return(idx - 1L)
      }

      fill_xml <- sprintf(
        "<fill xmlns=\"%s\"><patternFill patternType=\"solid\"><fgColor rgb=\"FF%s\"/></patternFill></fill>",
        private$ssml,
        sig
      )
      private$add_to_collection("d1:fills", fill_xml, sig, "fill_sigs")
    },

    #' @description Get or create a border index.
    get_border_id = function(
      top_style = NULL,
      top_color = NULL,
      bottom_style = NULL,
      bottom_color = NULL,
      left_style = NULL,
      left_color = NULL,
      right_style = NULL,
      right_color = NULL
    ) {
      sig <- paste(
        top_style %||% "",
        top_color %||% "",
        bottom_style %||% "",
        bottom_color %||% "",
        left_style %||% "",
        left_color %||% "",
        right_style %||% "",
        right_color %||% "",
        sep = "|"
      )
      idx <- match(sig, private$border_sigs)
      if (!is.na(idx)) {
        return(idx - 1L)
      }

      make_side <- function(tag, style, color) {
        if (is.null(style) || style == "") {
          return(sprintf("<%s/>", tag))
        }
        sprintf(
          "<%s style=\"%s\"><color rgb=\"FF%s\"/></%s>",
          tag,
          style,
          toupper(color %||% "000000"),
          tag
        )
      }
      left_xml <- make_side("left", left_style, left_color)
      right_xml <- make_side("right", right_style, right_color)
      top_xml <- make_side("top", top_style, top_color)
      bottom_xml <- make_side("bottom", bottom_style, bottom_color)

      border_xml <- sprintf(
        "<border xmlns=\"%s\">%s%s%s%s<diagonal/></border>",
        private$ssml,
        left_xml,
        right_xml,
        top_xml,
        bottom_xml
      )
      private$add_to_collection("d1:borders", border_xml, sig, "border_sigs")
    },

    #' @description Get or create a cell format (xf) index.
    get_style_id = function(
      font_id = 0L,
      fill_id = 0L,
      border_id = 0L,
      num_fmt_id = 0L,
      halign = NA,
      valign = NA,
      text_rotation = 0L,
      wrap_text = FALSE
    ) {
      sig <- paste(
        font_id,
        fill_id,
        border_id,
        num_fmt_id,
        halign %||% "",
        valign %||% "",
        text_rotation,
        as.integer(wrap_text),
        sep = "|"
      )
      idx <- match(sig, private$xf_sigs)
      if (!is.na(idx)) {
        return(idx - 1L)
      }

      apply_parts <- c(
        if (font_id > 0L) " applyFont=\"1\"",
        if (fill_id > 0L) " applyFill=\"1\"",
        if (border_id > 0L) " applyBorder=\"1\"",
        if (num_fmt_id > 0L) " applyNumberFormat=\"1\""
      )

      has_align <- (!is.na(halign) && halign != "") ||
        (!is.na(valign) && valign != "") ||
        text_rotation > 0L ||
        wrap_text
      if (has_align) {
        apply_parts <- c(apply_parts, " applyAlignment=\"1\"")
      }

      align_xml <- ""
      if (has_align) {
        align_parts <- c(
          if (!is.na(halign) && halign != "") {
            sprintf(" horizontal=\"%s\"", halign)
          },
          if (!is.na(valign) && valign != "") {
            sprintf(" vertical=\"%s\"", valign)
          },
          if (text_rotation > 0L) {
            sprintf(" textRotation=\"%d\"", text_rotation)
          },
          if (wrap_text) " wrapText=\"1\""
        )
        align_xml <- sprintf(
          "<alignment%s/>",
          paste0(align_parts, collapse = "")
        )
      }

      xf_xml <- sprintf(
        "<xf xmlns=\"%s\" numFmtId=\"%d\" fontId=\"%d\" fillId=\"%d\" borderId=\"%d\" xfId=\"0\"%s>%s</xf>",
        private$ssml,
        num_fmt_id,
        font_id,
        fill_id,
        border_id,
        paste0(apply_parts, collapse = ""),
        align_xml
      )
      private$add_to_collection("d1:cellXfs", xf_xml, sig, "xf_sigs")
    },

    #' @description Get or create a cell format index for a number format.
    get_xf_id = function(num_fmt_id) {
      self$get_style_id(num_fmt_id = as.integer(num_fmt_id))
    },

    #' @description Save styles.xml to disk.
    save = function() {
      write_xml(private$doc, private$file)
    }
  ),

  private = list(
    file = NULL,
    doc = NULL,
    ns = NULL,
    ssml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    font_sigs = character(0),
    fill_sigs = character(0),
    border_sigs = character(0),
    xf_sigs = character(0),

    index_existing = function() {
      ns <- private$ns

      # fonts
      font_nodes <- xml_find_all(private$doc, "d1:fonts/d1:font", ns = ns)
      private$font_sigs <- vapply(
        font_nodes,
        function(node) {
          nm <- xml_text(xml_find_first(node, "d1:name/@val", ns = ns))
          sz <- xml_text(xml_find_first(node, "d1:sz/@val", ns = ns))
          b <- !inherits(xml_find_first(node, "d1:b", ns = ns), "xml_missing")
          i <- !inherits(xml_find_first(node, "d1:i", ns = ns), "xml_missing")
          u <- !inherits(xml_find_first(node, "d1:u", ns = ns), "xml_missing")
          col_node <- xml_find_first(node, "d1:color", ns = ns)
          col <- if (inherits(col_node, "xml_missing")) {
            "000000"
          } else {
            rgb <- xml_attr(col_node, "rgb")
            if (!is.na(rgb) && nchar(rgb) == 8) substring(rgb, 3) else "theme"
          }
          paste(
            nm,
            sz,
            as.integer(b),
            as.integer(i),
            as.integer(u),
            col,
            sep = "|"
          )
        },
        character(1L)
      )

      # fills
      fill_nodes <- xml_find_all(private$doc, "d1:fills/d1:fill", ns = ns)
      private$fill_sigs <- vapply(
        fill_nodes,
        function(node) {
          fg <- xml_find_first(node, "d1:patternFill/d1:fgColor", ns = ns)
          if (inherits(fg, "xml_missing")) {
            return("none")
          }
          rgb <- xml_attr(fg, "rgb")
          if (!is.na(rgb) && nchar(rgb) == 8) {
            toupper(substring(rgb, 3))
          } else {
            "none"
          }
        },
        character(1L)
      )

      # borders
      border_nodes <- xml_find_all(
        private$doc,
        "d1:borders/d1:border",
        ns = ns
      )
      private$border_sigs <- vapply(
        border_nodes,
        function(node) {
          get_side <- function(side) {
            s <- xml_find_first(node, paste0("d1:", side), ns = ns)
            style <- xml_attr(s, "style")
            col_node <- xml_find_first(s, "d1:color", ns = ns)
            col <- if (inherits(col_node, "xml_missing")) {
              ""
            } else {
              rgb <- xml_attr(col_node, "rgb")
              if (!is.na(rgb) && nchar(rgb) == 8) substring(rgb, 3) else ""
            }
            paste(if (is.na(style)) "" else style, col, sep = "|")
          }
          paste(
            get_side("top"),
            get_side("bottom"),
            get_side("left"),
            get_side("right"),
            sep = "|"
          )
        },
        character(1L)
      )

      # cellXfs
      xf_nodes <- xml_find_all(private$doc, "d1:cellXfs/d1:xf", ns = ns)
      private$xf_sigs <- vapply(
        xf_nodes,
        function(node) {
          fid <- xml_attr(node, "fontId") %||% "0"
          filid <- xml_attr(node, "fillId") %||% "0"
          bid <- xml_attr(node, "borderId") %||% "0"
          nfid <- xml_attr(node, "numFmtId") %||% "0"
          align <- xml_find_first(node, "d1:alignment", ns = ns)
          ha <- if (inherits(align, "xml_missing")) {
            ""
          } else {
            xml_attr(align, "horizontal") %||% ""
          }
          va <- if (inherits(align, "xml_missing")) {
            ""
          } else {
            xml_attr(align, "vertical") %||% ""
          }
          rot <- if (inherits(align, "xml_missing")) {
            "0"
          } else {
            xml_attr(align, "textRotation") %||% "0"
          }
          wrap <- if (inherits(align, "xml_missing")) {
            "0"
          } else {
            xml_attr(align, "wrapText") %||% "0"
          }
          paste(fid, filid, bid, nfid, ha, va, rot, wrap, sep = "|")
        },
        character(1L)
      )
    },

    add_to_collection = function(xpath, xml_str, sig, sig_field) {
      parent <- xml_find_first(private$doc, xpath, ns = private$ns)
      xml_add_child(parent, read_xml(xml_str))
      count <- length(private[[sig_field]]) + 1L
      xml_attr(parent, "count") <- as.character(count)
      private[[sig_field]] <- c(private[[sig_field]], sig)
      count - 1L # 0-based
    }
  )
)

# validation helpers ----

validate_sheet_name <- function(label) {
  if (is.null(label) || !nzchar(label)) {
    cli::cli_abort("Sheet name cannot be empty.")
  }
  if (nchar(label) > 31L) {
    cli::cli_abort(
      "Sheet name {.val {label}} exceeds 31 characters ({nchar(label)})."
    )
  }
  if (grepl("[\\]\\[*?/\\\\:]", label, perl = TRUE)) {
    cli::cli_abort(
      "Sheet name {.val {label}} contains invalid characters."
    )
  }
  if (grepl("^'|'$", label)) {
    cli::cli_abort(
      "Sheet name {.val {label}} cannot start or end with a single quote."
    )
  }
  invisible(label)
}

check_sheet_exists <- function(x, sheet) {
  available <- x$worksheets$sheet_names()
  if (!sheet %in% available) {
    cli::cli_abort(c(
      "Sheet {.val {sheet}} not found.",
      "i" = "Available sheets: {.val {available}}."
    ))
  }
  invisible(sheet)
}

# read_xlsx ----
#' @export
#' @title Create an 'Excel' document object
#' @description Read and import an xlsx file as an R object
#' representing the document. This function is experimental.
#' @param path path to the xlsx file to use as base document.
#' @param x an rxlsx object
#' @examples
#' read_xlsx()
read_xlsx <- function(path = NULL) {
  if (!is.null(path) && !file.exists(path)) {
    cli::cli_abort("Could not find file {.file {path}}.")
  }

  if (is.null(path)) {
    path <- system.file(package = "officer", "template/template.xlsx")
  }

  if (!grepl("\\.xlsx$", path, ignore.case = TRUE)) {
    cli::cli_abort("{.fun read_xlsx} only supports {.file .xlsx} files.")
  }

  package_dir <- tempfile()
  unpack_folder(file = path, folder = package_dir)

  obj <- structure(
    list(package_dir = package_dir),
    .Names = c("package_dir"),
    class = "rxlsx"
  )

  obj$content_type <- content_type$new(package_dir)
  obj$worksheets <- worksheets$new(package_dir)
  obj$sheets <- dir_sheet$new(obj)
  obj$styles <- xlsx_styles$new(package_dir)
  obj$core_properties <- read_core_properties(obj$package_dir)

  obj
}

#' @export
#' @title Add a sheet
#' @description Add a sheet into an xlsx worksheet.
#' @param x rxlsx object
#' @param label sheet label
#' @examples
#' my_ws <- read_xlsx()
#' my_pres <- add_sheet(my_ws, label = "new sheet")
add_sheet <- function(x, label) {
  validate_sheet_name(label)
  if (label %in% x$worksheets$sheet_names()) {
    cli::cli_abort("Sheet {.val {label}} already exists.")
  }

  new_slidename <- x$worksheets$get_new_sheetname()

  xml_file <- file.path(x$package_dir, "xl/worksheets", new_slidename)

  template_file <- system.file(package = "officer", "template/sheet.xml")
  file.copy(template_file, xml_file, copy.mode = FALSE)

  rel_filename <- file.path(
    dirname(xml_file),
    "_rels",
    paste0(basename(xml_file), ".rels")
  )
  dir.create(dirname(rel_filename), showWarnings = FALSE)
  template_rel_file <- system.file(
    package = "officer",
    "template/sheet.xml.rels"
  )
  file.copy(template_rel_file, rel_filename, copy.mode = FALSE)

  # update presentation elements
  x$worksheets$add_sheet(
    target = file.path("worksheets", new_slidename),
    label = label
  )

  partname <- file.path("/xl/worksheets", new_slidename)
  override <- setNames(
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    partname
  )
  x$content_type$add_override(value = override)

  x$sheets$update()

  sheet_select(x, sheet = label)
}

#' @export
#' @rdname read_xlsx
length.rxlsx <- function(x) {
  x$sheets$length()
}

#' @export
#' @title Select sheet
#' @description Set a particular sheet selected when workbook will be
#' edited.
#' @param x rxlsx object
#' @param sheet sheet name
#' @examples
#' my_ws <- read_xlsx()
#' my_pres <- add_sheet(my_ws, label = "new sheet")
#' my_pres <- sheet_select(my_ws, sheet = "new sheet")
#' print(my_ws, target = tempfile(fileext = ".xlsx") )
sheet_select <- function(x, sheet) {
  check_sheet_exists(x, sheet)
  x$worksheets$view_on_sheet(sheet)
  active_index <- which(x$worksheets$sheet_names() == sheet)
  ns <- "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  for (i in seq_len(x$sheets$length())) {
    sheet_obj <- x$sheets$get_sheet(i)
    sv <- xml_find_first(
      sheet_obj$get(),
      "d1:sheetViews/d1:sheetView",
      ns = c(d1 = ns)
    )
    if (!inherits(sv, "xml_missing")) {
      if (i == active_index) {
        xml_attr(sv, "tabSelected") <- "1"
      } else {
        xml_attr(sv, "tabSelected") <- "0"
      }
    }
    sheet_obj$save()
  }
  x
}

#' @export
#' @title Write data to a sheet
#' @description Write a content into a sheet of an xlsx workbook.
#' Multiple calls can write to different positions on the same sheet.
#'
#' This is a generic function dispatching on `value`. Supported inputs:
#' - `data.frame`: written as a table (header in row `start_row`, data
#'   starting at `start_row + 1`).
#' - `character`: each element in its own cell. The `character` method
#'   accepts a `direction` argument (`"vertical"` -- default -- or
#'   `"horizontal"`) to stack elements in a column or a row.
#' - [fpar()]: richtext paragraph written into a single cell at
#'   `(start_row, start_col)`. Font, size, colour, bold, italic,
#'   underline, strikethrough and sub/superscript chunks are honoured.
#' - [block_list()]: one cell per `fpar` item. Also accepts
#'   `direction = "vertical"` (default) or `"horizontal"`.
#'
#' @param x rxlsx object
#' @param value a `data.frame`, `character` vector, [fpar()] or
#' [block_list()]
#' @param sheet sheet name (must already exist)
#' @param start_row row index where the header / first cell will be
#' written (default 1)
#' @param start_col column index where the first column / first cell
#' will be written (default 1)
#' @param ... method-specific arguments. In particular,
#' `direction = "vertical"` (default) or `"horizontal"` is honoured by
#' the `character` and `block_list` methods.
#' @example inst/examples/example-sheet_write_data.R
sheet_write_data <- function(x, value, sheet, start_row = 1L, start_col = 1L,
                             ...) {
  UseMethod("sheet_write_data", value)
}

#' @export
sheet_write_data.default <- function(x, value, sheet, ...) {
  cli::cli_abort(c(
    "No {.fun sheet_write_data} method for an object of class
     {.cls {class(value)[1]}}.",
    "i" = "Supported inputs: {.cls data.frame}, {.cls character},
           {.cls fpar}, {.cls block_list}."
  ))
}

#' @export
sheet_write_data.data.frame <- function(x, value, sheet,
                                        start_row = 1L, start_col = 1L, ...) {
  stopifnot(inherits(x, "rxlsx"))
  check_sheet_exists(x, sheet)
  start_row <- as.integer(start_row)
  start_col <- as.integer(start_col)

  # resolve style IDs for date/datetime columns
  date_xf <- NULL
  datetime_xf <- NULL
  for (j in seq_len(ncol(value))) {
    col <- value[[j]]
    if (inherits(col, "POSIXct") && is.null(datetime_xf)) {
      datetime_xf <- x$styles$get_xf_id(22L)
    } else if (inherits(col, "Date") && is.null(date_xf)) {
      date_xf <- x$styles$get_xf_id(14L)
    }
  }

  new_cells <- df_to_cells(
    value,
    start_row,
    start_col,
    date_xf = date_xf,
    datetime_xf = datetime_xf
  )

  write_cells_into_sheet(x, sheet, new_cells)
}

# Shared plumbing: merge a named list of row XML strings (keyed by row
# number as character) into the target sheet's <sheetData>.
write_cells_into_sheet <- function(x, sheet, new_cells) {
  sheet_obj <- x$sheets$get_sheet(
    which(x$worksheets$sheet_names() == sheet)
  )
  sheet_doc <- sheet_obj$get()
  ns <- c(d1 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  existing_rows <- xml_find_all(sheet_doc, "d1:sheetData/d1:row", ns = ns)

  if (length(existing_rows) == 0L) {
    # fast path: no existing data, skip merge
    row_nums <- as.integer(names(new_cells))
    idx <- order(row_nums)
    row_strs <- sprintf(
      "<row r=\"%d\">%s</row>",
      row_nums[idx],
      unlist(new_cells[idx])
    )
  } else {
    # read existing rows
    existing_map <- list()
    for (row_node in existing_rows) {
      r <- xml_attr(row_node, "r")
      cell_nodes <- xml_find_all(row_node, "d1:c", ns = ns)
      cell_strs <- vapply(cell_nodes, as.character, character(1L))
      existing_map[[r]] <- paste0(cell_strs, collapse = "")
    }

    # merge: new rows overwrite or extend existing rows
    for (r in names(new_cells)) {
      if (is.null(existing_map[[r]])) {
        existing_map[[r]] <- new_cells[[r]]
      } else {
        old_doc <- read_xml(paste0("<row>", existing_map[[r]], "</row>"))
        new_doc <- read_xml(paste0("<row>", new_cells[[r]], "</row>"))
        old_cells <- xml_find_all(old_doc, "c")
        new_cells_nodes <- xml_find_all(new_doc, "c")
        old_refs <- xml_attr(old_cells, "r")
        new_refs <- xml_attr(new_cells_nodes, "r")
        keep <- !old_refs %in% new_refs
        parts <- c(
          vapply(old_cells[keep], as.character, character(1L)),
          vapply(new_cells_nodes, as.character, character(1L))
        )
        existing_map[[r]] <- paste0(parts, collapse = "")
      }
    }

    row_nums <- as.integer(names(existing_map))
    existing_map <- existing_map[order(row_nums)]
    row_nums <- sort(row_nums)
    row_strs <- sprintf(
      "<row r=\"%d\">%s</row>",
      row_nums,
      unlist(existing_map)
    )
  }

  # update sheetData in the existing DOM (preserves namespaces, etc.)
  sheet_data_node <- xml_find_first(sheet_doc, "d1:sheetData", ns = ns)
  new_sheet_data <- read_xml(paste0(
    "<sheetData xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    paste0(row_strs, collapse = ""),
    "</sheetData>"
  ))
  xml_replace(sheet_data_node, new_sheet_data)

  # remove dimension (Excel recalculates it)
  dim_node <- xml_find_first(sheet_doc, "d1:dimension", ns = ns)
  if (!inherits(dim_node, "xml_missing")) {
    xml_remove(dim_node)
  }

  sheet_file <- sheet_obj$file_name()
  write_xml(sheet_doc, sheet_file)
  sheet_obj$feed(sheet_file)

  x
}

# Convert a data.frame to a named list of cell XML strings, keyed by row number
df_to_cells <- function(
  data,
  start_row = 1L,
  start_col = 1L,
  date_xf = NULL,
  datetime_xf = NULL
) {
  nr <- nrow(data)
  nc <- ncol(data)
  col_indices <- seq(start_col, length.out = nc)
  col_letters <- int_to_col(col_indices)
  header_row <- start_row
  data_rows <- seq(start_row + 1L, length.out = nr)

  # header cells
  header_cells <- sprintf(
    "<c r=\"%s%d\" t=\"inlineStr\"><is><t>%s</t></is></c>",
    col_letters,
    header_row,
    htmlEscapeCopy(names(data))
  )

  # data cells: matrix (nr x nc)
  cell_matrix <- matrix(NA_character_, nrow = nr, ncol = nc)
  for (j in seq_len(nc)) {
    col <- data[[j]]
    refs <- sprintf("%s%d", col_letters[j], data_rows)
    if (inherits(col, "POSIXct") && !is.null(datetime_xf)) {
      num_val <- sprintf("%.10f", (as.numeric(col) / 86400) + 25569)
      cells <- sprintf(
        "<c r=\"%s\" s=\"%d\"><v>%s</v></c>",
        refs,
        datetime_xf,
        num_val
      )
      cells[is.na(col)] <- sprintf("<c r=\"%s\"/>", refs[is.na(col)])
    } else if (inherits(col, "Date") && !is.null(date_xf)) {
      num_val <- sprintf("%.17G", as.numeric(col) + 25569)
      cells <- sprintf(
        "<c r=\"%s\" s=\"%d\"><v>%s</v></c>",
        refs,
        date_xf,
        num_val
      )
      cells[is.na(col)] <- sprintf("<c r=\"%s\"/>", refs[is.na(col)])
    } else if (is.logical(col)) {
      cells <- sprintf(
        "<c r=\"%s\" t=\"b\"><v>%d</v></c>",
        refs,
        as.integer(col)
      )
      cells[is.na(col)] <- sprintf("<c r=\"%s\"/>", refs[is.na(col)])
    } else if (is.numeric(col)) {
      val_str <- sprintf("%.17G", col)
      cells <- sprintf("<c r=\"%s\"><v>%s</v></c>", refs, val_str)
      cells[is.na(col)] <- sprintf("<c r=\"%s\"/>", refs[is.na(col)])
    } else {
      col <- as.character(col)
      na_mask <- is.na(col)
      cells <- sprintf(
        "<c r=\"%s\" t=\"inlineStr\"><is><t>%s</t></is></c>",
        refs,
        htmlEscapeCopy(col)
      )
      cells[na_mask] <- sprintf("<c r=\"%s\"/>", refs[na_mask])
    }
    cell_matrix[, j] <- cells
  }

  # build named list: row number -> already-collapsed row strings
  row_cells <- do.call(paste0, as.data.frame(cell_matrix))
  result <- setNames(
    as.list(c(paste0(header_cells, collapse = ""), row_cells)),
    as.character(c(header_row, data_rows))
  )
  result
}

# --- Richtext cell helpers ------------------------------------------------
# Convert one `ftext`-style chunk (list with $value and $pr = fp_text)
# to a spreadsheetml <r><rPr>...</rPr><t xml:space="preserve">...</t></r>.
chunk_to_xlsx_run <- function(chunk) {
  pr <- chunk$pr
  rpr_parts <- character(0)
  if (isTRUE(pr$bold)) rpr_parts <- c(rpr_parts, "<b/>")
  if (isTRUE(pr$italic)) rpr_parts <- c(rpr_parts, "<i/>")
  if (isTRUE(pr$underlined)) rpr_parts <- c(rpr_parts, "<u/>")
  if (isTRUE(pr$strike)) rpr_parts <- c(rpr_parts, "<strike/>")
  if (!is.na(pr$font.size)) {
    rpr_parts <- c(rpr_parts,
                   sprintf("<sz val=\"%g\"/>", pr$font.size))
  }
  if (!is.na(pr$color) && !identical(pr$color, "transparent")) {
    rpr_parts <- c(rpr_parts,
                   sprintf("<color rgb=\"FF%s\"/>",
                           hex_color(pr$color)))
  }
  if (!is.na(pr$font.family)) {
    rpr_parts <- c(rpr_parts,
                   sprintf("<rFont val=\"%s\"/>",
                           htmlEscapeCopy(pr$font.family)))
  }
  if (identical(pr$vertical.align, "superscript")) {
    rpr_parts <- c(rpr_parts, "<vertAlign val=\"superscript\"/>")
  } else if (identical(pr$vertical.align, "subscript")) {
    rpr_parts <- c(rpr_parts, "<vertAlign val=\"subscript\"/>")
  }

  rpr <- if (length(rpr_parts)) {
    paste0("<rPr>", paste0(rpr_parts, collapse = ""), "</rPr>")
  } else {
    ""
  }

  sprintf(
    "<r>%s<t xml:space=\"preserve\">%s</t></r>",
    rpr,
    htmlEscapeCopy(chunk$value)
  )
}

# Convert a fpar to the <is>...</is> body of an inline-string cell.
fpar_to_inlinestr <- function(fpar) {
  chunks <- fpar$chunks
  if (!length(chunks)) {
    return("<is><t xml:space=\"preserve\"/></is>")
  }
  runs <- vapply(chunks, chunk_to_xlsx_run, character(1L))
  paste0("<is>", paste0(runs, collapse = ""), "</is>")
}

# Build the full <c r="..."> cell for a fpar at a given ref.
fpar_to_cell_xml <- function(fpar, ref) {
  sprintf(
    "<c r=\"%s\" t=\"inlineStr\">%s</c>",
    ref,
    fpar_to_inlinestr(fpar)
  )
}

# --- sheet_write_data methods: character, fpar, block_list ----------------

#' @export
sheet_write_data.character <- function(x, value, sheet,
                                       start_row = 1L, start_col = 1L,
                                       direction = c("vertical",
                                                     "horizontal"),
                                       ...) {
  stopifnot(inherits(x, "rxlsx"))
  check_sheet_exists(x, sheet)
  direction <- match.arg(direction)
  start_row <- as.integer(start_row)
  start_col <- as.integer(start_col)

  n <- length(value)
  if (direction == "vertical") {
    rows <- seq(start_row, length.out = n)
    cols <- rep(start_col, n)
  } else {
    rows <- rep(start_row, n)
    cols <- seq(start_col, length.out = n)
  }
  refs <- sprintf("%s%d", int_to_col(cols), rows)

  na_mask <- is.na(value)
  cells <- character(n)
  cells[!na_mask] <- sprintf(
    "<c r=\"%s\" t=\"inlineStr\"><is><t xml:space=\"preserve\">%s</t></is></c>",
    refs[!na_mask],
    htmlEscapeCopy(value[!na_mask])
  )
  cells[na_mask] <- sprintf("<c r=\"%s\"/>", refs[na_mask])

  # group by row
  new_cells <- tapply(cells, rows, paste0, collapse = "")
  new_cells <- as.list(new_cells)
  names(new_cells) <- as.character(unique(rows))

  write_cells_into_sheet(x, sheet, new_cells)
}

#' @export
sheet_write_data.fpar <- function(x, value, sheet,
                                  start_row = 1L, start_col = 1L, ...) {
  stopifnot(inherits(x, "rxlsx"))
  check_sheet_exists(x, sheet)
  start_row <- as.integer(start_row)
  start_col <- as.integer(start_col)

  ref <- sprintf("%s%d", int_to_col(start_col), start_row)
  cell <- fpar_to_cell_xml(value, ref)

  new_cells <- setNames(list(cell), as.character(start_row))
  write_cells_into_sheet(x, sheet, new_cells)
}

#' @export
sheet_write_data.block_list <- function(x, value, sheet,
                                        start_row = 1L, start_col = 1L,
                                        direction = c("vertical",
                                                      "horizontal"),
                                        ...) {
  stopifnot(inherits(x, "rxlsx"))
  check_sheet_exists(x, sheet)
  direction <- match.arg(direction)
  start_row <- as.integer(start_row)
  start_col <- as.integer(start_col)

  # only fpar items are rendered -- skip non-fpar blocks with a warning
  is_fpar <- vapply(value, inherits, logical(1), what = "fpar")
  if (!all(is_fpar)) {
    cli::cli_warn(c(
      "{.fun sheet_write_data} can only render {.cls fpar} items in a
       {.cls block_list}; skipping {sum(!is_fpar)} non-fpar item{?s}."
    ))
  }
  fpars <- value[is_fpar]
  n <- length(fpars)
  if (n == 0L) return(invisible(x))

  if (direction == "vertical") {
    rows <- seq(start_row, length.out = n)
    cols <- rep(start_col, n)
  } else {
    rows <- rep(start_row, n)
    cols <- seq(start_col, length.out = n)
  }
  refs <- sprintf("%s%d", int_to_col(cols), rows)

  cells <- vapply(seq_len(n), function(i) {
    fpar_to_cell_xml(fpars[[i]], refs[i])
  }, character(1L))

  new_cells <- tapply(cells, rows, paste0, collapse = "")
  new_cells <- as.list(new_cells)
  names(new_cells) <- as.character(unique(rows))

  write_cells_into_sheet(x, sheet, new_cells)
}

#' @export
#' @title Add a drawing to an Excel sheet
#' @description Add a graphical element into a sheet of an xlsx workbook.
#' This is a generic function dispatching on `value`. Methods are
#' provided by extension packages 'mschart' and 'rvg'.
#'
#' Use [sheet_write_data()] to write data into the sheet before or
#' after adding a drawing.
#' @param x rxlsx object created by [read_xlsx()]
#' @param value object to add (dispatched to the appropriate method)
#' @param sheet sheet name (must already exist, see [add_sheet()])
#' @param ... additional arguments passed to methods
#' @return the rxlsx object (invisibly)
#' @seealso [read_xlsx()], [add_sheet()], [sheet_write_data()]
sheet_add_drawing <- function(x, value, sheet, ...) {
  UseMethod("sheet_add_drawing", value)
}

#' @export
#' @method sheet_add_drawing external_img
#' @title Add an image to an Excel sheet
#' @description Add an image file (PNG, JPEG, GIF, ...) to a sheet in
#'   an xlsx workbook created with [read_xlsx()]. The image is copied
#'   into `xl/media/` and placed on the sheet via an absolute anchor
#'   (inch-based position and size).
#' @param x rxlsx object (created by [read_xlsx()])
#' @param value an [external_img()] object
#' @param sheet sheet name (must already exist)
#' @param left,top top-left anchor of the image, in inches. Defaults
#'   to `(1, 1)`.
#' @param width,height size of the image, in inches. When `NULL`
#'   (default) the dimensions stored on `value` (from `external_img()`)
#'   are used.
#' @param ... unused
#' @return the rxlsx object (invisibly)
#' @examples
#' img <- system.file("extdata", "example.png", package = "officer")
#' if (nzchar(img) && file.exists(img)) {
#'   x <- read_xlsx()
#'   x <- add_sheet(x, label = "pics")
#'   x <- sheet_add_drawing(
#'     x, sheet = "pics",
#'     value = external_img(img, width = 2, height = 2)
#'   )
#'   print(x, target = tempfile(fileext = ".xlsx"))
#' }
#' @seealso [sheet_add_drawing()], [external_img()]
sheet_add_drawing.external_img <- function(x, value, sheet,
                                           left = 1, top = 1,
                                           width = NULL, height = NULL,
                                           ...) {
  stopifnot(inherits(x, "rxlsx"))
  check_sheet_exists(x, sheet)

  dims <- attr(value, "dims")
  if (is.null(width))  width  <- dims$width[1]
  if (is.null(height)) height <- dims$height[1]
  alt <- attr(value, "alt") %||% ""

  src <- as.character(value)[1]
  if (!file.exists(src)) {
    cli::cli_abort("Image file {.path {src}} does not exist.")
  }

  # copy image to xl/media/ with a unique name (file.copy overwrites
  # if the destination already exists)
  media_dir <- file.path(x$package_dir, "xl", "media")
  dir.create(media_dir, showWarnings = FALSE, recursive = TRUE)
  ext <- tolower(tools::file_ext(src))
  if (!nzchar(ext)) ext <- "png"
  target_path <- tempfile(pattern = "img",
                          tmpdir = media_dir,
                          fileext = paste0(".", ext))
  file.copy(src, target_path, overwrite = TRUE)

  sheet_obj <- x$sheets$get_sheet(
    which(x$worksheets$sheet_names() == sheet)
  )
  drawing <- xlsx_drawing$new(x$package_dir, sheet_obj, x$content_type)

  image_rid <- drawing$add_image_rel(basename(target_path))
  drawing$add_image_anchor(
    image_rid = image_rid,
    left = left, top = top,
    width = width, height = height,
    alt = alt
  )

  invisible(x)
}

# Convert integer column index to Excel letter (1=A, 2=B, ..., 27=AA)
int_to_col <- function(x) {
  vapply(
    x,
    function(n) {
      if (is.na(n)) {
        return(NA_character_)
      }
      col <- ""
      while (n > 0) {
        n <- n - 1L
        col <- paste0(LETTERS[n %% 26L + 1L], col)
        n <- n %/% 26L
      }
      col
    },
    character(1L)
  )
}

#' @export
#' @param target path to the xlsx file to write
#' @param ... unused
#' @rdname read_xlsx
#' @examples
#' x <- read_xlsx()
#' print(x, target = tempfile(fileext = ".xlsx"))
print.rxlsx <- function(x, target = NULL, ...) {
  if (is.null(target)) {
    cat("xlsx document with", length(x), "sheet(s):\n")
    print(x$worksheets$sheet_names())
    return(invisible())
  }

  if (!grepl(x = target, pattern = "\\.(xlsx)$", ignore.case = TRUE)) {
    cli::cli_abort("{.file {target}} should have {.file .xlsx} extension.")
  }

  x$worksheets$save()
  x$content_type$save()
  x$styles$save()

  x$core_properties['modified', 'value'] <- format(
    Sys.time(),
    "%Y-%m-%dT%H:%M:%SZ"
  )
  x$core_properties['lastModifiedBy', 'value'] <- Sys.getenv("USER")
  write_core_properties(x$core_properties, x$package_dir)

  invisible(pack_folder(folder = x$package_dir, target = target))
}
