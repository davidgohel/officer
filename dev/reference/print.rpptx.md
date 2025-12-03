# Write a 'PowerPoint' file.

Create a 'PowerPoint' file from an `rpptx` object (created by
[`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md)).

## Usage

``` r
# S3 method for class 'rpptx'
print(x, target = NULL, preview = FALSE, ...)
```

## Arguments

- x:

  an `rpptx` object.

- target:

  path to the .pptx file to write. If `target` is `NULL` (default), the
  `rpptx` object is printed to the console.

- preview:

  Save `x` to a temporary file and open it (default `FALSE`).

- ...:

  unused.

## Value

If preview is `TRUE`, returns the temp file path invisibly.

## See also

[`read_pptx()`](https://davidgohel.github.io/officer/dev/reference/read_pptx.md)

## Examples

``` r
# write an rpptx object to a .pptx file ----
file <- tempfile(fileext = ".pptx")
x <- read_pptx() # empty presentation, has no slides yet
print(x, target = file)

# preview mode: save to temp file and open locally ----
if (FALSE) { # \dontrun{
print(x, preview = TRUE)
} # }
```
