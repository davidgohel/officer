# Convert Data URIs to PNG Files

Decodes base64-encoded data URIs and writes them to PNG files.

## Usage

``` r
base64_to_image(data_uri, output_files)
```

## Arguments

- data_uri:

  Character, a data URI character vector starting with
  "data:image/png;base64,"

- output_files:

  Character, paths to the output PNG files

## Value

Character, the paths to the created PNG files

## Examples

``` r
rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
base64_str <- image_to_base64(rlogo)
base64_to_image(
  data_uri = base64_str,
  output_files = tempfile(fileext = ".jpeg")
)
#> [1] "/tmp/RtmpeBjOjU/file18605c282253.jpeg"
```
