# Get the dimensions of an image

Get the width and height in pixels of an image file or of a raster
object.

The dimensions of 'png', 'jpeg' and 'gif' files are read from their
header, without decoding the image, which is much faster. Other formats
are read with package 'magick', which must then be installed.

## Usage

``` r
image_dims(x)
```

## Arguments

- x:

  path to an image file, or a raster object (see
  [`grDevices::as.raster()`](https://rdrr.io/r/grDevices/as.raster.html)).

## Value

A named numeric vector: `width` and `height`, in pixels.

## Examples

``` r
img <- file.path(R.home("doc"), "html", "logo.jpg")
image_dims(img)
#>  width height 
#>    100     76 

image_dims(as.raster(matrix(0:1, nrow = 2, ncol = 3)))
#>  width height 
#>      3      2 
```
