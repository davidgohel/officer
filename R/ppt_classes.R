# content_type -----
content_type <- R6Class(
  "content_type",
  public = list(

    initialize = function( x ) {

      private$filename <- file.path(x$package_dir, "[Content_Types].xml")

      doc <- read_xml(x = private$filename )
      ns <- xml_ns(doc)

      extension <- xml_find_all(doc, "//*[contains(local-name(), 'Default')]/@Extension", ns = ns)
      extension <- xml_text( extension )
      content_type <- xml_find_all(doc, "//*[contains(local-name(), 'Default')]/@ContentType", ns = ns)
      content_type <- xml_text( content_type )
      names(content_type) <- extension
      private$default <- content_type

      partname <- xml_find_all(doc, "//*[contains(local-name(), 'Override')]/@PartName", ns = ns)
      partname <- xml_text(partname)
      content_type <- xml_find_all(doc, "//*[contains(local-name(), 'Override')]/@ContentType", ns = ns)
      content_type <- xml_text(content_type)
      names(content_type) <- partname
      private$override <- content_type

    },

    add_slide = function(partname){
      partname <- basename(partname)
      partname <- file.path("/ppt", "slides", partname )
      content_type <- setNames("application/vnd.openxmlformats-officedocument.presentationml.slide+xml", partname )
      override <- c( private$override, content_type )
      private$override <- override
      self
    },
    remove_slide = function(partname){
      id <- which( basename(names(private$override)) %in% basename(partname) )
      private$override <- private$override[-id]
      self
    },
    add_ext = function( extension, type ){
      if( !type %in% private$default && !extension %in% names(private$default) ){
        content_type <- setNames(type, extension )
        default <- c( private$default, content_type )
        private$default <- default
      }
      self
    },

    save = function() {
      self$add_ext(extension = "jpeg", type = "image/jpeg")
      self$add_ext(extension = "png", type = "image/png")
      self$add_ext(extension = "emf", type = "image/x-emf")
      self$add_ext(extension = "jpg", type = "application/octet-stream")
      attribs <- attr_chunk(c(xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"))
      out <- paste0("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                    "\n<Types", attribs, ">")
      if(length(private$default) > 0 ){
        default <- sprintf("<Default Extension=\"%s\" ContentType=\"%s\"/>", names(private$default), private$default )
        default <- paste0(default, collapse = "")
        out <- paste0(out, default )
      }
      if(length(private$override) > 0 ){
        override <- sprintf("<Override PartName=\"%s\" ContentType=\"%s\"/>", names(private$override), private$override )
        override <- paste0(override, collapse = "")
        out <- paste0(out, override )
      }
      out <- paste0(out, "</Types>" )
      cat(out, file = private$filename)
      self

    }
  ),
  private = list(
    filename = NULL,
    default = NULL,
    override = NULL
  )
)

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
      tibble(id = id, rid = rid)
    }
  )
)


# openxml_document --------------------------------------------------------
openxml_document <- R6Class(
  "openxml_document",
  public = list(

    initialize = function( dir ) {
      private$reldir = dir
    },

    feed = function( file ) {
      private$filename <- file
      private$rels_filename <- file.path( dirname(file), "_rels", paste0(basename(file), ".rels") )

      private$doc <- read_xml(file)
      private$rels_doc <- relationship$new()$feed_from_xml(private$rels_filename)
      self
    },
    file_name = function(){
      private$filename
    },
    name = function(){
      basename(private$filename)
    },
    get = function(){
      private$doc
    },
    dir_name = function(){
      private$reldir
    },
    save = function() {
      write_xml(private$doc, file = private$filename)
      private$rels_doc$write(private$rels_filename)
      self
    },
    remove = function() {
      unlink(private$filename)
      unlink(private$rels_filename)
      private$filename
    },
    rel_df = function(){
      private$rels_doc$get_data()
    },
    relationship = function(){
      private$rels_doc
    }

  ),
  private = list(

    filename = NULL,
    rels_filename = NULL,
    doc = NULL,
    rels_doc = NULL,
    reldir = NULL

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
      nodeset <- xml_find_all( self$get(), "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]")
      read_xfrm(nodeset, self$file_name(), self$name())
    }


  ),
  private = list(

    theme_file = function(){
      data <- self$rel_df()
      theme_file <- data[basename(data$type) == "theme", "target", drop = TRUE]
      file.path( "ppt/theme", basename(theme_file) )
    }

  )

)

# slide_layout ------------------------------------------------------------

#' @importFrom dplyr group_by_
#' @importFrom dplyr mutate_
#' @importFrom dplyr ungroup
slide_layout <- R6Class(
  "slide_layout",
  inherit = openxml_document,
  public = list(

    get_metadata = function( ){
      rels <- self$rel_df()
      rels <- rels[basename( rels$type ) == "slideMaster", ]
      tibble(name = self$name(), filename = self$file_name(), master_file = rels$target)
    },
    xfrm = function(){
      rels <- self$rel_df()
      rels <- rels[basename( rels$type ) == "slideMaster", ]

      nodeset <- xml_find_all( self$get(), "p:cSld/p:spTree/*[self::p:sp or self::p:graphicFrame or self::p:grpSp or self::p:pic]")
      data <- read_xfrm(nodeset, self$file_name(), self$name())
      data <- group_by_(data, .dots = c("id", "type"))
      data <- mutate_(data, num = "row_number()")
      data <- ungroup(data)
      data$master_file <- basename(rels$target)
      data
    },
    name = function(){
      xmldoc <- read_xml(self$file_name())
      csld <- xml_find_first(xmldoc, "//p:cSld")
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
      filter_criteria <- interp(~ basename(type) == "slideLayout")
      slide_info <- filter_(private$rels_doc$get_data() , filter_criteria)
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
      self$save()
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
      # <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide3.xml"/>
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
          id <- which( out$type == type )[index]
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
      tibble(name = self$name(), filename = self$file_name(), layout_file = rels$target)
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
      private$collection <- map( filenames, function(x, container){
        container$clone()$feed(x)
      }, container = container)

    },
    get_metadata = function(){
      map_df(private$collection, function(x) x$get_metadata())
    },
    names = function(){
      map_chr(private$collection, function(x) x$name())
    },
    xfrm = function( ){
      map_df(private$collection, function(x) x$xfrm() )
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
      unames <- map_chr(private$collection, function(x) x$name())
      ufnames <- map_chr(private$collection, function(x) x$file_name())
      tibble(master_name = unames, filename = ufnames)
    },
    get_color_scheme = function( ){
      map_df(private$collection, function(x) x$colors())
    }

  )
)


# dir_layout ---------------------------------------------------------
#' @importFrom dplyr inner_join
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
      out <- inner_join(data_layouts, data_masters, by = "master_file")
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
#' @importFrom dplyr between
dir_slide <- R6Class(
  "dir_slide",
  inherit = dir_collection,
  public = list(

    initialize = function( x ) {
      super$initialize(x, slide$new("ppt/slides"))
      private$slides_path <- file.path(x$package_dir, "ppt/slides")
      private$slide_layouts <- dir_layout$new( x )
      map(private$collection, function(x, ref) x$set_xfrm(ref), ref = private$slide_layouts$get_xfrm_data() )
    },

    update = function( ){
      dir_ <- file.path(private$package_dir, "ppt/slides")
      filenames <- list.files(path = dir_, pattern = "\\.xml$", full.names = TRUE)
      private$collection <- map( filenames, function(x, container){
        container$clone()$feed(x)$fortify_id()
      }, container = slide$new("ppt/slides"))
      private$slides_path <- file.path(private$package_dir, "ppt/slides")
      private$collection <- map(private$collection, function(x, ref) x$set_xfrm(ref), ref = private$slide_layouts$get_xfrm_data() )
      self
    },
    remove = function(index ){
      slide_obj <- private$collection[[index]]
      private$collection <- private$collection[-index]
      slide_obj$remove()
    },
    get_xfrm = function( ){
      map(private$collection, function(x, ref) x$get_xfrm() )
    },


    get_slide = function(id){
      l_ <- self$length()
      if( is.null(id) || !between(id, 1, l_ ) ){
        stop("unvalid id ", id, " (", l_," slide(s))", call. = FALSE)
      }
      private$collection[[id]]
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

      slide_files <- basename( list.files(slide_dir, pattern = "\\.xml$") )
      slidename <- "slide1.xml"
      if( length(slide_files)){
        slide_index <- as.integer(gsub("^(slide)([0-9]+)(\\.xml)$", "\\2", slide_files ))
        slidename <- gsub(pattern = "[0-9]+", replacement = max(slide_index) + 1, slidename)
      }
      slidename
    },
    layout_files = function(){
      map_chr(private$collection, function(x) x$layout_name())
    }
  ),
  private = list(
    slides_path = NULL,
    slide_layouts = NULL
  )
)

