# PowerPoint table to matrix

Convert the data in an a 'PowerPoint' table to a matrix or all data to a
list of matrices.

## Usage

``` r
# S3 method for class 'rpptx'
as.matrix(
  x,
  ...,
  slide_id = NA_integer_,
  id = NA_character_,
  span = c(NA_character_, "fill")
)
```

## Arguments

- x:

  The rpptx object to convert (as created by
  [`read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.md))

- ...:

  Ignored

- slide_id:

  The slide number to load from (NA indicates first slide with a table,
  NULL indicates all slides and all tables)

- id:

  The table ID to load from (ignored it `is.null(slide_id)`, NA
  indicates to load the first table from the `slide_id`)

- span:

  How should col_span/row_span values be handled? `NA` means to leave
  the value as `NA`, and `"fill"` means to fill matrix cells with the
  value.

## Value

A matrix with the data, or if `slide_id=NULL`, a list of matrices

## Examples

``` r
library(officer)
pptx_file <- system.file(package="officer", "doc_examples", "example.pptx")
z <- read_pptx(pptx_file)
as.matrix(z, slide_id = NULL)
#> $`1`
#> $`1`$`18`
#>      [,1]        [,2]       [,3]            
#> [1,] "Header 1 " "Header 2" "Header 3"      
#> [2,] "A"         "12.23"    "blah blah"     
#> [3,] "B"         "1.23"     "blah blah blah"
#> [4,] "B"         "9.0"      "Salut"         
#> [5,] "C"         "6"        "Hello"         
#> 
#> 
#> $`3`
#> $`3`$`5`
#>      [,1] [,2] [,3] [,4] [,5]
#> [1,] ""   ""   ""   ""   ""  
#> [2,] "a"  NA   ""   ""   "d" 
#> [3,] ""   ""   "c"  ""   NA  
#> [4,] "b"  ""   NA   ""   ""  
#> [5,] NA   ""   ""   ""   ""  
#> 
#> 
```
