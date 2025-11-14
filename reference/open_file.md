# Opens a file locally

Opening a file locally requires a compatible application to be installed
(e.g., MS Office or LibreOffice for .pptx or .docx files).

## Usage

``` r
open_file(path)
```

## Arguments

- path:

  File path.

## Details

*NB:* Function is a small wrapper around
[`utils::browseURL()`](https://rdrr.io/r/utils/browseURL.html) to have a
more suitable function name.

## Examples

``` r
x <- read_pptx()
x <- add_slide(x, "Title Slide", ctrTitle = "My Title")
file <- print(x, tempfile(fileext = ".pptx"))
if (FALSE) { # \dontrun{
open_file(file)} # }
```
