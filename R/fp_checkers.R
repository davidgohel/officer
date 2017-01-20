check_spread_integer <- function( obj, value, dest){
  varname <- as.character(substitute(value))
  if( is.numeric( value ) && length(value) == 1  && value >= 0 ){
    for(i in dest)
      obj[[i]] <- as.integer(value)
  } else stop(varname, " must be a positive integer scalar.", call. = FALSE)
  obj
}

check_set_color <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !is.color( value ) )
    stop(varname, " must be a valid color.", call. = FALSE )
  else obj[[varname]] <- value
  obj
}

check_set_pic <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !grepl(pattern = "^rId[0-9]+", value) )
    stop(varname, " must be a valid reference id: ", value, call. = FALSE )
  obj[[varname]] <- value
  obj
}
check_set_file <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !file.exists(value) )
    stop(varname, " must be a valid filename.", call. = FALSE )
  obj[[varname]] <- value
  obj
}
check_set_border <- function( obj, value){
  varname <- as.character(substitute(value))
  if( !inherits( value, "fp_border" ) )
    stop(varname, " must be a fp_border object." , call. = FALSE)
  else obj[[varname]] <- value
  obj
}

check_spread_border <- function( obj, value, dest ){
  varname <- as.character(substitute(value))
  if( !inherits( value, "fp_border" ) )
    stop(varname, " must be a fp_border object." , call. = FALSE)
  for(i in dest )
    obj[[i]] <- value
  obj
}

check_set_integer <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.numeric( value ) && length(value) == 1  && value >= 0 ){
    obj[[varname]] <- as.integer(value)
  } else stop(varname, " must be a positive integer scalar.", call. = FALSE)
  obj
}

check_set_bool <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.logical( value ) && length(value) == 1 ){
    obj[[varname]] <- value
  } else stop(varname, " must be a boolean", call. = FALSE)
  obj
}
check_set_chr <- function( obj, value){
  varname <- as.character(substitute(value))
  if( is.character( value ) && length(value) == 1 ){
    obj[[varname]] <- value
  } else stop(varname, " must be a string", call. = FALSE)
  obj
}

check_set_choice <- function( obj, value, choices){
  varname <- as.character(substitute(value))
  if( is.character( value ) && length(value) == 1 ){
    if( !value %in% choices )
      stop(varname, " must be one of ", paste( shQuote(choices), collapse = ", "), call. = FALSE )
    obj[[varname]] = value
  } else stop(varname, " must be a character scalar.", call. = FALSE)
  obj
}

