# presentation ------------------------------------------------------------

presentation <- R6Class(
  "presentation",
  inherit = openxml_document,

  public = list(

    initialize = function( x ) {
      super$initialize(character(0))
      presentation_filename <- file.path(x$package_dir, "ppt", "presentation.xml")
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

      xml_list <- xml_find_first(private$doc, "//p:sldIdLst")
      xml_elt <- paste(
        sprintf("<p:sldId id=\"%.0f\" r:id=\"%s\"/>", private$slide_id, private$slide_rid),
        collapse = "" )
      xml_elt <- paste0( pml_with_ns("p:sldIdLst"),  xml_elt, "</p:sldIdLst>")
      xml_elt <- as_xml_document(xml_elt)

      if( !inherits(xml_list, "xml_missing")){
        xml_replace(xml_list, xml_elt)
      } else{ ## needs to be after sldMasterIdLst...
        xml_add_sibling(xml_find_first(private$doc, "//p:sldMasterIdLst"), xml_elt)
      }

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

      xml_list <- xml_find_first(private$doc, "//p:sldIdLst")
      xml_elt <- paste(
        sprintf("<p:sldId id=\"%.0f\" r:id=\"%s\"/>", private$slide_id, private$slide_rid),
        collapse = "" )

      xml_elt <- paste0(pml_with_ns("p:sldIdLst"), xml_elt, "</p:sldIdLst>")
      xml_elt <- as_xml_document(xml_elt)

      if( !inherits(xml_list, "xml_missing")){
        xml_replace(xml_list, xml_elt)
      } else{ ## needs to be after sldMasterIdLst...
        xml_add_sibling(xml_find_first(private$doc, "//p:sldMasterIdLst"), xml_elt)
      }

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

    set_xfrm = function(xfrm_ref){
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


# dir_collection ---------------------------------------------------------

dir_collection <- R6Class(
  "dir_collection",
  public = list(

    initialize = function( x, container ) {
      private$package_dir <- x$package_dir
      dir_ <- file.path(private$package_dir, container$dir_name())
      filenames <- list.files(path = dir_, pattern = "\\.xml$", full.names = TRUE)
      private$collection <- lapply( filenames, function(x, container){
        container$clone()$feed(x)
      }, container = container)
      names(private$collection) <- basename(filenames)
    },

    collection_get = function(name){
      private$collection[[name]]
    },

    get_metadata = function(){
      dat <- lapply(private$collection, function(x) x$get_metadata())
      rbind.match.columns(dat)
    },
    names = function(){
      sapply(private$collection, function(x) x$name())
    },
    xfrm = function( ){
      dat <- lapply(private$collection, function(x) x$xfrm() )
      rbind.match.columns(dat)
    }
  ),

  private = list(

    collection = NULL,
    package_dir = NULL

  )
)




# dir_master ---------------------------------------------------------

dir_master <- R6Class(
  "dir_master",
  inherit = dir_collection,
  public = list(

    get_metadata = function( ){
      unames <- sapply(private$collection, function(x) x$name())
      ufnames <- sapply(private$collection, function(x) x$file_name())
      data.frame(stringsAsFactors = FALSE, master_name = unames, filename = ufnames)
    },
    get_color_scheme = function( ){
      dat <- lapply(private$collection, function(x) x$colors())
      rbind.match.columns(dat)
    }

  )
)


# dir_layout ---------------------------------------------------------
dir_layout <- R6Class(
  "dir_layout",
  inherit = dir_collection,
  public = list(
    initialize = function( x ) {
      super$initialize(x, slide_layout$new("ppt/slideLayouts"))
      private$master_collection <- dir_master$new(x, slide_master$new("ppt/slideMasters") )
      private$xfrm_data <- xfrmize(self$xfrm(), private$master_collection$xfrm())
    },

    get_xfrm_data = function(){
      private$xfrm_data
    },

    get_metadata = function( ){
      data_layouts <- super$get_metadata()
      data_masters <- private$master_collection$get_metadata()
      data_masters$master_file <- basename(data_masters$filename)
      data_masters$filename <- NULL
      data_layouts$master_file <- basename(data_layouts$master_file)
      out <- merge(data_layouts, data_masters, by = "master_file", all = FALSE)
      out$filename <- basename(out$filename)
      out
    },

    get_master = function(){
      private$master_collection
    }

  ),
  private = list(
    master_collection = NULL,
    xfrm_data = NULL
  )
)


# dir_slide ---------------------------------------------------------
dir_slide <- R6Class(
  "dir_slide",
  inherit = dir_collection,
  public = list(

    initialize = function( x ) {
      super$initialize(x, slide$new("ppt/slides"))
      private$slides_path <- file.path(x$package_dir, "ppt/slides")
      private$slide_layouts <- dir_layout$new( x )
      private$collection <- lapply(private$collection, function(x, ref) x$set_xfrm(ref), ref = private$slide_layouts$get_xfrm_data() )
      names(private$collection) <- sapply(private$collection, function(x) x$name() )
      private$slides_list <- private$get_slide_list()

    },

    add_slide = function(slide_file){

      slide <- slide$new("ppt/slides")
      slide$feed(slide_file)
      slide$set_xfrm(private$slide_layouts$get_xfrm_data())

      collect <- private$collection
      new_elt <- list(slide)
      names(new_elt) <- basename(slide_file)
      collect <- append(collect, new_elt)

      sl_id <- as.integer( gsub( "(slide)([0-9]+)(\\.xml)$", "\\2", names(collect) ) )
      private$collection <- collect[order(sl_id)]
      private$slides_list <- names(private$collection)

      self
    },

    remove_slide = function(index ){
      slide_obj <- private$collection[[index]]
      private$collection <- private$collection[-index]
      private$slides_list <- names(private$collection)
      slide_obj$remove()
    },

    save_slides = function(){
      lapply( private$collection, function(x){
        x$save()
      } )
      self
    },

    get_xfrm = function( ){
      lapply(private$collection, function(x) x$get_xfrm() )
    },


    get_slide = function(id){
      l_ <- self$length()
      if( is.null(id) || !between(id, 1, l_ ) ){
        stop("unvalid id ", id, " (", l_," slide(s))", call. = FALSE)
      }
      index <- which( names(private$collection) == private$slides_list[id])
      private$collection[[index]]
    },

    get_metadata = function(){
      super$get_metadata()
    },

    length = function(){
      length(private$collection)
    },

    get_new_slidename = function(){
      slide_dir <- file.path(private$package_dir, "ppt/slides")
      if( !file.exists(slide_dir)){
        dir.create(file.path(slide_dir, "_rels"), showWarnings = FALSE, recursive = TRUE)
      }

      slide_files <- names(private$collection)
      slidename <- "slide1.xml"
      if( length(slide_files)){
        slide_index <- as.integer(gsub("^(slide)([0-9]+)(\\.xml)$", "\\2", slide_files ))
        slidename <- gsub(pattern = "[0-9]+", replacement = max(slide_index) + 1, slidename)
      }
      slidename
    },
    layout_files = function(){
      sapply(private$collection, function(x) x$layout_name())
    }
  ),
  private = list(
    slides_path = NULL,
    slides_list = NULL,
    slide_layouts = NULL,

    get_slide_list = function(){
      slide_dir <- file.path(private$package_dir, "ppt/slides")
      slide_files <- list.files(slide_dir, pattern = "\\.xml$")
      slide_index <- seq_along(slide_files)
      if( length(slide_files)){
        slide_files <- basename( slide_files )
        slide_index <- as.integer(gsub("^(slide)([0-9]+)(\\.xml)$", "\\2", slide_files ))
        slide_files <- slide_files[order(slide_index)]
      }
      slide_files
    }


  )
)

