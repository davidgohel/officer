# presentation ------------------------------------------------------------

presentation <- R6Class(
  "presentation",
  inherit = openxml_document,

  public = list(

    initialize = function( package_dir ) {
      super$initialize(character(0))
      presentation_filename <- file.path(package_dir, "ppt", "presentation.xml")
      self$feed(presentation_filename)

      slide_df <- private$get_slide_df()
      private$slide_id <- slide_df$id
      private$slide_rid <- slide_df$rid

    },

    add_slide = function(target){


      private$rels_doc$add(id = paste0("rId", private$rels_doc$get_next_id() ),
                           type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                           target = target )
      rels <- private$rels_doc$get_data()
      rid <- rels[rels$target %in% target,"id"]


      ids <- private$slide_id
      if( length( ids ) < 1 )
        new_id <- 256
      else new_id <- max(ids) + 1

      private$slide_id <- c( private$slide_id, new_id)
      private$slide_rid <- c( private$slide_rid, rid)

      private$update_xml()

      self
    },
    slide_data = function(){
      rel_df <- self$rel_df()
      rel_df <- rel_df[, c("id", "target")]
      names(rel_df) <- c("slide_rid", "target")
      ref <- data.frame(slide_id = private$slide_id,
                 slide_rid = private$slide_rid,
                 stringsAsFactors = FALSE)
      base::merge(x = ref, y = rel_df, sort = FALSE,
                  by = "slide_rid", all.x = TRUE, all.y = FALSE)
    },

    move_slide = function(from, to){
      slide_list <- self$slide_data()

      ord <- seq_len(nrow(slide_list))
      if( from < to ) to <- to+1
      ord[ord >= to] <- ord[ord >= to] + 1L
      ord[from] <- to
      slide_list <- slide_list[order(ord),,drop = FALSE]
      private$slide_id <- slide_list$slide_id
      private$slide_rid <- slide_list$slide_rid

      private$update_xml()

      self
    },

    remove_slide = function(target){

      reldf <- self$rel_df()
      id <- which( basename(reldf$target) %in% basename(target)  )
      rid <- reldf$id[id]
      private$rels_doc$remove(target = target )

      dropid <- which( private$slide_rid %in% rid )

      private$slide_id <- private$slide_id[-dropid]
      private$slide_rid <- private$slide_rid[-dropid]

      private$update_xml()

      self
    }

  ),
  private = list(

    slide_id = NULL,
    slide_rid = NULL,

    get_slide_df = function() {
      nodes <- xml_find_all(private$doc, "//p:sldIdLst/p:sldId")
      id <- as.integer( xml_attr(nodes, "id", ns = xml_ns(private$doc)) )
      rid <- xml_attr(nodes, "r:id", ns = xml_ns(private$doc))
      data.frame(id = id, rid = rid, stringsAsFactors = FALSE)
    },
    update_xml = function(){
      xml_list <- xml_find_first(private$doc, "//p:sldIdLst")
      xml_elt <- paste(
        sprintf("<p:sldId id=\"%.0f\" r:id=\"%s\"/>", private$slide_id, private$slide_rid),
        collapse = "" )

      xml_elt <- paste0(pml_with_ns("p:sldIdLst"), xml_elt, "</p:sldIdLst>")
      xml_elt <- as_xml_document(xml_elt)

      if( !inherits(xml_list, "xml_missing")){
        xml_replace(xml_list, xml_elt)
      } else{ ## needs to be after all MasterIdLst elements. placing it before sldSz seems to be the safest option.
        xml_add_sibling(xml_find_first(private$doc, "//p:sldSz"), xml_elt, .where = "before")
      }

      self
    }
  )
)




# slide master ------------------------------------------------------------
#' @importFrom xml2 xml_child
slide_master <- R6Class(
  "slide_master",
  inherit = openxml_document,
  public = list(

    name = function(){
      theme_ <- private$theme_file()
      root <- gsub( paste0(self$dir_name(), "$"), "", dirname( private$filename ) )
      xml_attr(read_xml(file.path( root,theme_)), "name")
    },

    colors = function(){
      theme_ <- private$theme_file()
      root <- gsub( paste0(self$dir_name(), "$"), "", dirname( private$filename ) )

      doc <- read_xml(file.path( root,theme_))
      read_theme_colors( doc, self$name() )
    },

    xfrm = function(){
      nodeset <- xml_find_all( self$get(), as_xpath_content_sel("p:cSld/p:spTree/") )
      read_xfrm(nodeset, self$file_name(), self$name())
    }


  ),
  private = list(

    theme_file = function(){
      data <- self$rel_df()
      theme_file <- data$target[basename(data$type) == "theme"]
      file.path( "ppt/theme", basename(theme_file) )
    }

  )

)

# slide_layout ------------------------------------------------------------
slide_layout <- R6Class(
  "slide_layout",
  inherit = openxml_document,
  public = list(

    get_metadata = function( ){
      rels <- self$rel_df()
      rels <- rels[basename( rels$type ) == "slideMaster", ]
      data.frame(stringsAsFactors = FALSE, name = self$name(), filename = self$file_name(), master_file = rels$target)
    },
    xfrm = function(){
      rels <- self$rel_df()
      rels <- rels[basename( rels$type ) == "slideMaster", ]

      nodeset <- xml_find_all( self$get(), as_xpath_content_sel("p:cSld/p:spTree/"))
      data <- read_xfrm(nodeset, self$file_name(), self$name())
      if( nrow(data))
        data$master_file <- basename(rels$target)
      else data$master_file <- character(0)
      data
    },
    write_template = function(new_file){

      path <- system.file(package = "officer", "template/slide.xml")

      rel_filename <- file.path(
        dirname(new_file), "_rels",
        paste0(basename(new_file), ".rels") )

      newrel <- relationship$new()$add(
        id = "rId1", type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        target = file.path("../slideLayouts", basename(self$file_name())) )
      newrel$write(path = rel_filename)
      file.copy(path, to = new_file)
      self
    },

    name = function(){
      csld <- xml_find_first(self$get(), "//p:cSld")
      xml_attr(csld, "name")
    }


  )

)

# slide ------------------------------------------------------------
slide <- R6Class(
  "slide",
  inherit = openxml_document,
  public = list(

    feed = function( file ) {
      super$feed(file)
      slide_info <- private$rels_doc$get_data()
      slide_info <- slide_info[basename(slide_info$type) == "slideLayout",]
      private$layout_file <- basename( slide_info$target )
      self
    },

    set_layout_xfrm = function(xfrm_ref){
      private$element_data <- xfrm_ref[xfrm_ref$file == private$layout_file,]
      self
    },
    fortify_id = function(){
      cnvpr <- xml_find_all(private$doc, "//p:cNvPr")
      for(i in seq_along(cnvpr))
        xml_attr( cnvpr[[i]], "id") <- i
      self
    },

    reference_img = function(src, dir_name){
      src <- unique( src )
      private$rels_doc$add_img(src, root_target = "../media")
      dir.create(dir_name, recursive = TRUE, showWarnings = FALSE)
      file.copy(from = src, to = file.path(dir_name, basename(src)))
      self
    },

    reference_slide = function(slide_file){

      rel_dat <- private$rels_doc$get_data()

      if( !slide_file %in% rel_dat$target ){
        next_id <- private$rels_doc$get_next_id()
        private$rels_doc$add(id = sprintf("rId%.0f", next_id),
                             type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                             target = slide_file )
      }
      self
    },

    reference_hyperlink = function(href){

      rel_dat <- private$rels_doc$get_data()

      if( !href %in% rel_dat$target ){
        next_id <- private$rels_doc$get_next_id()
        private$rels_doc$add(id = sprintf("rId%.0f", next_id),
                             type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                             target = href, target_mode = "External" )
      }
      # <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide3.xml"/>
      self
    },

    get_xfrm = function(type = NULL, index = 1){
      out <- private$element_data
      if( !is.null(type) ){
        if( type %in% out$type ){
          type_matches <- which( out$type == type )
          if( index <= length(type_matches) )
            id <- type_matches[index]
          else stop(type, " can only have ", length(type_matches), " element(s) but index is set to ", index)

        } else stop("type ", type, " is not available in the slide layout")
        out <- out[id, ]
      }
      out
    },
    get_location = function(type = NULL, index = 1){
      out <- private$element_data
      if( !is.null(type) ){
        if( type %in% out$type ){
          type_matches <- which( out$type == type )
          if( index <= length(type_matches) )
            id <- type_matches[index]
          else stop(type, " can only have ", length(type_matches), " element(s) but index is set to ", index)

        } else stop("type ", type, " is not available in the slide layout")
        out <- out[id, ]
      }

      out[c("offx", "offy", "cx", "cy")] <- lapply( out[c("offx", "offy", "cx", "cy")], function(x) x / 914400 )
      out <- out[c("offx", "offy", "cx", "cy", "ph_label", "type", "ph")]
      names(out) <- c("left", "top", "width", "height", "ph_label", "type", "ph")
      out
    },

    layout_name = function(){
      private$layout_file
    },

    get_metadata = function( ){
      rels <- self$rel_df()
      rels <- rels[basename( rels$type ) == "slideLayout", ]
      data.frame(stringsAsFactors = FALSE, name = self$name(), filename = self$file_name(), layout_file = rels$target)
    }

  ),
  private = list(
    layout_file = NULL,
    element_data = NULL
  )

)



