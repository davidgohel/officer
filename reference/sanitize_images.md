# Remove unused media from a document

The function will scan the media directory and delete images that are
not used anymore. This function is to be used when images have been
replaced many times.

## Usage

``` r
sanitize_images(x, warn_user = TRUE)
```

## Arguments

- x:

  `rdocx` or `rpptx` object

- warn_user:

  TRUE to make sure users are warned when using this function that will
  be un-exported in a next version
