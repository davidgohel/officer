# body_pr -----
#' @export
body_pr <- R6Class(
  "body_pr",
  public = list(

    initialize = function() {
      private$preset_text_warp <- any_content$new()
      private$autofit <- any_content$new()
      private$scene3d <- any_content$new()
      private$scene3d_sp <- any_content$new()
      private$ext_list <- any_content$new()
      private$attrs <- character(0)
    },
    set_attr = function(name, value){
      private$attrs[name] <- value
      self
    },
    set_autofit = function(autofit = TRUE){
      if( autofit )
        private$autofit <- any_content$new()$set_content("<a:spAutoFit/>")
      else private$autofit <- NULL
      self
    },
    get_attr = function(name){
      private$attrs[name]
    },
    feed_from_node = function(node, ns){
      private$attrs <- xml_attrs(node, ns )

      preset_text_warp <- xml_find_first(x = node, xpath = paste0(xml_path(node), "/a:prstTxWarp"), ns = ns )
      if( !inherits(preset_text_warp, "xml_missing") )
        private$preset_text_warp <- as.character(preset_text_warp)

      autofit_path <- paste0( xml_path(node), "/*[contains(local-name(), '",
                              c('noAutofit', 'normAutofit', 'spAutoFit'),
                              "')]", collapse = " | " )
      autofit <- xml_find_first(x = node, xpath = autofit_path, ns = ns )
      if( !inherits(autofit, "xml_missing") )
        private$autofit <- as.character(autofit)

      scene3d <- xml_find_first(x = node, xpath = paste0(xml_path(node), "/a:scene3d"), ns = ns )
      if( !inherits(scene3d, "xml_missing") )
        private$scene3d <- as.character(scene3d)

      scene3d_sp_path <- paste0( xml_path(node), "/*[contains(local-name(), '",
                                 c('sp3d', 'flatTx '), "')]", collapse = " | " )
      scene3d_sp <- xml_find_first(x = node, xpath = scene3d_sp_path, ns = ns )
      if( !inherits(scene3d_sp, "xml_missing") )
        private$scene3d_sp <- as.character(scene3d_sp)

      ext_list <- xml_find_first(nodelist, xpath = "//a:extLst", ns = ns )
      if( !inherits(ext_list, "xml_missing") )
        private$ext_list <- any_content$new(ext_list)

      self
    },

    format = function() {
      attribs <- attr_chunk(private$attrs)
      paste0("<a:bodyPr", attribs, ">",
             ifelse( is.null( private$preset_text_warp ), "", private$preset_text_warp$format() ),
             ifelse( is.null( private$autofit ), "", private$autofit$format() ),
             ifelse( is.null( private$scene3d ), "", private$scene3d$format() ),
             ifelse( is.null( private$scene3d_sp ), "", private$scene3d_sp$format() ),
             ifelse( is.null( private$ext_list ), "", private$ext_list$format() ),
             "</a:bodyPr>")
    }
  ),
  private = list(
    attrs = NULL,
    preset_text_warp = NULL,
    autofit = NULL,
    scene3d = NULL,
    scene3d_sp = NULL,
    ext_list = NULL
  )
)





