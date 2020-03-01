#' @export
#' @title number of slides
#' @description Function \code{length} will return the number of slides.
#' @param x an rpptx object
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres)
#' my_pres <- add_slide(my_pres)
#' length(my_pres)
#' @family functions for reading presentation informations
length.rpptx <- function( x ){
  x$slide$length()
}

#' @export
#' @title slides width and height
#' @description get the width and height of slides in inches as
#' a named vector.
#' @inheritParams length.rpptx
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres,
#'   layout = "Two Content", master = "Office Theme")
#' slide_size(my_pres)
#' @family functions for reading presentation informations
slide_size <- function(x) {
  pres <- x$presentation$get()
  dimensions <- xml_attrs(xml_find_first(pres, "p:sldSz"))
  dimensions <- as.list(as.integer(dimensions[c("cx", "cy")]) / 914400)
  names(dimensions) <- c("width", "height")
  dimensions
}



#' @export
#' @title presentation layouts summary
#' @description get informations about slide layouts and
#' master layouts into a data.frame. This function returns
#' a data.frame containing all layout and master names.
#' @inheritParams length.rpptx
#' @examples
#' my_pres <- read_pptx()
#' layout_summary ( x = my_pres )
#' @family functions for reading presentation informations
layout_summary <- function( x ){
  data <- x$slideLayouts$get_metadata()
  data.frame(layout = data$name, master = data$master_name, stringsAsFactors = FALSE)
}

#' @export
#' @title slide layout properties
#' @description get information about a particular slide layout
#' into a data.frame.
#' @inheritParams length.rpptx
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @examples
#' x <- read_pptx()
#' layout_properties ( x = x, layout = "Title Slide", master = "Office Theme" )
#' layout_properties ( x = x, master = "Office Theme" )
#' layout_properties ( x = x, layout = "Two Content" )
#' layout_properties ( x = x )
#' @family functions for reading presentation informations
layout_properties <- function( x, layout = NULL, master = NULL ){

  data <- x$slideLayouts$get_xfrm_data()

  if( !is.null(layout) && !is.null(master) ){
    data <- data[data$name == layout & data$master_name == master,]
  } else if( is.null(layout) && !is.null(master) ){
    data <- data[data$master_name == master,]
  } else if( !is.null(layout) && is.null(master) ){
    data <- data[data$name == layout,]
  }
  data <- data[,c("master_name", "name", "type", "id", "ph_label", "ph", "offx", "offy", "cx", "cy")]
  data[["offx"]] <- data[["offx"]] / 914400
  data[["offy"]] <- data[["offy"]] / 914400
  data[["cx"]] <- data[["cx"]] / 914400
  data[["cy"]] <- data[["cy"]] / 914400

  data
}

#' @export
#' @title PowerPoint placeholder parameters annotation
#' @description generates a slide from each layout in the base document to
#' identify the placeholder indexes, types, names, master names and layout names.
#'
#' This is to be used when need to know what parameters should be used with
#' \code{ph_location*} calls. The parameters are printed in their corresponding shapes.
#'
#' Note that if there are duplicated \code{ph_label}, you should not use \code{ph_location_label}.
#'
#' @param path path to the pptx file to use as base document or NULL to use the officer default
#' @param output_file filename to store the annotated powerpoint file or NULL to suppress generation
#' @return rpptx object of the annotated PowerPoint file
#' @examples
#' # To generate an anotation of the default base document with officer:
#' annotate_base(output_file = tempfile(fileext = ".pptx"))
#'
#' # To generate an annotation of the base document 'mydoc.pptx' and place the
#' # annotated output in 'mydoc_annotate.pptx'
#' # annotate_base(path = 'mydoc.pptx', output_file='mydoc_annotate.pptx')
#'
#' @family functions for reading presentation informations
annotate_base <- function(path = NULL, output_file = 'annotated_layout.pptx' ){

  ppt <- read_pptx(path=path)

  # Pulling out all of the layouts stored in the template
  lay_sum <- layout_summary(ppt)

  # Looping through each layout

  for(lidx in seq_len(nrow(lay_sum))){
    # Pulling out the layout properties

    layout <- lay_sum[lidx, 1]
    master <- lay_sum[lidx, 2]
    lp <- layout_properties ( x = ppt, layout = layout, master = master)
    lp <- lp[order(lp$type, as.integer(lp$id)),]

    id <- unlist( lapply(rle( lp$type)$lengths, function(x){
      seq_len(x)
    }) )
    textstr <- sprintf('type="%s", index =%d, ph_label="%s"',
                       lp$type, id, lp$ph_label)

    # Adding a slide for the current layout
    ppt <- add_slide(x=ppt, layout = layout, master = master)
    size <- slide_size(ppt)
    fpar_ <- fpar(sprintf('layout ="%s", master = "%s"', layout, master),
                  fp_t = fp_text(color = "orange", font.size = 20),
                  fp_p = fp_par(text.align = "right", padding = 5)
    )
    ppt <- ph_with(x = ppt, value = fpar_, ph_label = "layout_ph",
                   location = ph_location(left = 0, top = -0.5, width = size$width, height = 1,
                               bg = "transparent", newlabel = "layout_ph"))


    # Blank slides have nothing
    if(length(lp[,1] > 0)){
      # Now we go through each placholder
      for(pidx in seq_len(nrow(lp))){
        ppt <- ph_with(x=ppt, value = textstr[pidx], location = ph_location_type(type = lp$type[pidx], id = id[pidx]))
      }
    }
  }

  if(!is.null(output_file)){
    print(ppt, target = output_file)
  }

  ppt
}

#' @export
#' @title get PowerPoint slide content in a data.frame
#' @description get content and positions of current slide
#' into a data.frame. Data for any tables, images, or paragraphs are
#' imported into the resulting data.frame.
#' @note
#' The column \code{id} of the result is not to be used by users.
#' This is a technical string id whose value will be used by office
#' when the document will be rendered. This is not related to argument
#' \code{index} required by functions \code{ph_with}.
#' @inheritParams length.rpptx
#' @param index slide index
#' @examples
#' my_pres <- read_pptx()
#' my_pres <- add_slide(my_pres)
#' my_pres <- ph_with(my_pres, format(Sys.Date()),
#'   location = ph_location_type(type="dt"))
#' my_pres <- add_slide(my_pres)
#' my_pres <- ph_with(my_pres, iris[1:2,],
#'   location = ph_location_type(type="body"))
#' slide_summary(my_pres)
#' slide_summary(my_pres, index = 1)
#' @family functions for reading presentation informations
slide_summary <- function( x, index = NULL ){

  l_ <- length(x)
  if( l_ < 1 ){
    stop("presentation contains no slide", call. = FALSE)
  }

  if( is.null(index) )
    index <- x$cursor

  if( !between(index, 1, l_ ) ){
    stop("unvalid index ", index, " (", l_," slide(s))", call. = FALSE)
  }

  slide <- x$slide$get_slide(index)

  nodes <- xml_find_all(slide$get(), as_xpath_content_sel("p:cSld/p:spTree/") )
  data <- read_xfrm(nodes, file = "slide", name = "" )
  data$text <- sapply(nodes, xml_text )
  data[["offx"]] <- data[["offx"]] / 914400
  data[["offy"]] <- data[["offy"]] / 914400
  data[["cx"]] <- data[["cx"]] / 914400
  data[["cy"]] <- data[["cy"]] / 914400

  data$name <- NULL
  data$file <- NULL
  data$ph <- NULL
  data
}






#' @export
#' @title color scheme
#' @description get master layout color scheme into a data.frame.
#' @inheritParams length.rpptx
#' @examples
#' x <- read_pptx()
#' color_scheme ( x = x )
#' @family functions for reading presentation informations
color_scheme <- function( x ){
  x$masterLayouts$get_color_scheme()
}
