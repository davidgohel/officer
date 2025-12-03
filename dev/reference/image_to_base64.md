# Images to base64

encodes images into base64 strings.

## Usage

``` r
image_to_base64(filepaths)
```

## Arguments

- filepaths:

  file names.

## Examples

``` r
rlogo <- file.path(R.home("doc"), "html", "logo.jpg")
base64_str <- image_to_base64(rlogo)
```
