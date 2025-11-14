# Extract media from a document object

Extract files from a `rpptx` object.

## Usage

``` r
media_extract(x, path, target)
```

## Arguments

- x:

  an rpptx object

- path:

  media path, should be a relative path

- target:

  target file

## Examples

``` r
example_pptx <- system.file(package = "officer",
  "doc_examples/example.pptx")
doc <- read_pptx(example_pptx)
content <- pptx_summary(doc)
image_row <- content[content$content_type %in% "image", ]
media_file <- image_row$media_file
png_file <- tempfile(fileext = ".png")
media_extract(doc, path = media_file, target = png_file)
#> [1] TRUE
```
