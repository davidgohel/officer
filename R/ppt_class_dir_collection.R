# dir_collection ---------------------------------------------------------
dir_collection <- R6Class(
  "dir_collection",
  public = list(

    initialize = function( package_dir, container ) {
      dir_ <- file.path(package_dir, container$dir_name())
      private$package_dir <- package_dir

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


# dir_layout ---------------------------------------------------------
dir_layout <- R6Class(
  "dir_layout",
  inherit = dir_collection,
  public = list(
    initialize = function( package_dir, master_metadata, master_xfrm ) {
      super$initialize(package_dir, slide_layout$new("ppt/slideLayouts"))
      private$master_metadata <- master_metadata
      private$xfrm_data <- xfrmize(self$xfrm(), master_xfrm)
    },

    get_xfrm_data = function(){
      private$xfrm_data
    },

    get_metadata = function( ){
      data_layouts <- super$get_metadata()
      data_masters <- private$master_metadata
      data_masters$master_file <- basename(data_masters$filename)
      data_masters$filename <- NULL
      data_layouts$master_file <- basename(data_layouts$master_file)
      out <- merge(data_layouts, data_masters, by = "master_file", all = FALSE)
      out$filename <- basename(out$filename)
      out
    }

  ),
  private = list(
    master_collection = NULL,
    master_metadata = NULL,
    xfrm_data = NULL
  )
)


# dir_slide ---------------------------------------------------------
dir_slide <- R6Class(
  "dir_slide",
  inherit = dir_collection,
  public = list(

    initialize = function( package_dir, xfrm_layout_data ) {
      super$initialize(package_dir, slide$new("ppt/slides"))
      private$collection <- lapply(private$collection, function(x, ref) x$set_layout_xfrm(ref), ref = xfrm_layout_data )
      names(private$collection) <- sapply(private$collection, function(x) x$name() )
      private$slides_list <- private$get_slide_list()

    },

    add_slide = function(slide_file, xfrm_layout_data){

      slide <- slide$new("ppt/slides")
      slide$feed(slide_file)
      slide$set_layout_xfrm(xfrm_layout_data)

      collect <- private$collection
      new_elt <- list(slide)
      names(new_elt) <- basename(slide_file)
      collect <- append(collect, new_elt)

      sl_id <- as.integer( gsub( "(slide)([0-9]+)(\\.xml)$", "\\2", names(collect) ) )
      private$collection <- collect[order(sl_id)]
      private$slides_list <- names(private$collection)

      self
    },
    slide_index = function( name ){
      which( private$slides_list %in% name )
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
    }
  ),
  private = list(
    slides_list = NULL,

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

