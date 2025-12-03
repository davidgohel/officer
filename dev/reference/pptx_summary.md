# PowerPoint content in a data.frame

Read content of a PowerPoint document and return a dataset representing
the document.

## Usage

``` r
pptx_summary(x, preserve = FALSE)
```

## Arguments

- x:

  an rpptx object

- preserve:

  If `FALSE` (default), text in table cells is collapsed into a single
  line. If `TRUE`, line breaks in table cells are preserved as a "\n"
  character. This feature is adapted from
  `docxtractr::docx_extract_tbl()` published under a [MIT
  licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE)
  in the 'docxtractr' package by Bob Rudis.

## Examples

``` r
example_pptx <- system.file(package = "officer",
  "doc_examples/example.pptx")
doc <- read_pptx(example_pptx)
pptx_summary(doc)
#>                    text id content_type slide_id row_id cell_id col_span
#> 1                 Title 12    paragraph        1     NA      NA       NA
#> 2              A table  13    paragraph        1     NA      NA       NA
#> 3         and some text 13    paragraph        1     NA      NA       NA
#> 4     and some list (1) 13    paragraph        1     NA      NA       NA
#> 5     and some list (2) 13    paragraph        1     NA      NA       NA
#> 1.1           Header 1  18   table cell        1      1       1        1
#> 1.4                   A 18   table cell        1      2       1        1
#> 1.7                   B 18   table cell        1      3       1        1
#> 1.10                  B 18   table cell        1      4       1        1
#> 1.13                  C 18   table cell        1      5       1        1
#> 2.2            Header 2 18   table cell        1      1       2        1
#> 2.5               12.23 18   table cell        1      2       2        1
#> 2.8                1.23 18   table cell        1      3       2        1
#> 2.11                9.0 18   table cell        1      4       2        1
#> 2.14                  6 18   table cell        1      5       2        1
#> 3.3            Header 3 18   table cell        1      1       3        1
#> 3.6           blah blah 18   table cell        1      2       3        1
#> 3.9      blah blah blah 18   table cell        1      3       3        1
#> 3.12              Salut 18   table cell        1      4       3        1
#> 3.15              Hello 18   table cell        1      5       3        1
#> 11               R logo 15    paragraph        1     NA      NA       NA
#> 12                 <NA> 17        image        1     NA      NA       NA
#> 13                   Hi  2    paragraph        2     NA      NA       NA
#> 21             This is   3    paragraph        2     NA      NA       NA
#> 31         an unordered  3    paragraph        2     NA      NA       NA
#> 41   list of paragraphs  3    paragraph        2     NA      NA       NA
#> 51                       3    paragraph        2     NA      NA       NA
#> 6    This is an ordered  3    paragraph        2     NA      NA       NA
#> 7    list of paragraphs  3    paragraph        2     NA      NA       NA
#> 1.12                     5   table cell        3      1       1        1
#> 1.6                   a  5   table cell        3      2       1        2
#> 1.11                     5   table cell        3      3       1        1
#> 1.16                  b  5   table cell        3      4       1        1
#> 1.21               <NA>  5   table cell        3      5       1        1
#> 2.21                     5   table cell        3      1       2        1
#> 2.7                <NA>  5   table cell        3      2       2        0
#> 2.12                     5   table cell        3      3       2        1
#> 2.17                     5   table cell        3      4       2        1
#> 2.22                     5   table cell        3      5       2        1
#> 3.31                     5   table cell        3      1       3        1
#> 3.8                      5   table cell        3      2       3        1
#> 3.13                  c  5   table cell        3      3       3        1
#> 3.18               <NA>  5   table cell        3      4       3        1
#> 3.23                     5   table cell        3      5       3        1
#> 4.4                      5   table cell        3      1       4        1
#> 4.9                      5   table cell        3      2       4        1
#> 4.14                     5   table cell        3      3       4        1
#> 4.19                     5   table cell        3      4       4        1
#> 4.24                     5   table cell        3      5       4        1
#> 5.5                      5   table cell        3      1       5        1
#> 5.10                  d  5   table cell        3      2       5        1
#> 5.15               <NA>  5   table cell        3      3       5        1
#> 5.20                     5   table cell        3      4       5        1
#> 5.25                     5   table cell        3      5       5        1
#>      row_span            media_file
#> 1          NA                  <NA>
#> 2          NA                  <NA>
#> 3          NA                  <NA>
#> 4          NA                  <NA>
#> 5          NA                  <NA>
#> 1.1         1                  <NA>
#> 1.4         1                  <NA>
#> 1.7         1                  <NA>
#> 1.10        1                  <NA>
#> 1.13        1                  <NA>
#> 2.2         1                  <NA>
#> 2.5         1                  <NA>
#> 2.8         1                  <NA>
#> 2.11        1                  <NA>
#> 2.14        1                  <NA>
#> 3.3         1                  <NA>
#> 3.6         1                  <NA>
#> 3.9         1                  <NA>
#> 3.12        1                  <NA>
#> 3.15        1                  <NA>
#> 11         NA                  <NA>
#> 12         NA ppt/media//image1.png
#> 13         NA                  <NA>
#> 21         NA                  <NA>
#> 31         NA                  <NA>
#> 41         NA                  <NA>
#> 51         NA                  <NA>
#> 6          NA                  <NA>
#> 7          NA                  <NA>
#> 1.12        1                  <NA>
#> 1.6         1                  <NA>
#> 1.11        1                  <NA>
#> 1.16        2                  <NA>
#> 1.21        0                  <NA>
#> 2.21        1                  <NA>
#> 2.7         1                  <NA>
#> 2.12        1                  <NA>
#> 2.17        1                  <NA>
#> 2.22        1                  <NA>
#> 3.31        1                  <NA>
#> 3.8         1                  <NA>
#> 3.13        2                  <NA>
#> 3.18        0                  <NA>
#> 3.23        1                  <NA>
#> 4.4         1                  <NA>
#> 4.9         1                  <NA>
#> 4.14        1                  <NA>
#> 4.19        1                  <NA>
#> 4.24        1                  <NA>
#> 5.5         1                  <NA>
#> 5.10        2                  <NA>
#> 5.15        0                  <NA>
#> 5.20        1                  <NA>
#> 5.25        1                  <NA>
pptx_summary(example_pptx)
#>                    text id content_type slide_id row_id cell_id col_span
#> 1                 Title 12    paragraph        1     NA      NA       NA
#> 2              A table  13    paragraph        1     NA      NA       NA
#> 3         and some text 13    paragraph        1     NA      NA       NA
#> 4     and some list (1) 13    paragraph        1     NA      NA       NA
#> 5     and some list (2) 13    paragraph        1     NA      NA       NA
#> 1.1           Header 1  18   table cell        1      1       1        1
#> 1.4                   A 18   table cell        1      2       1        1
#> 1.7                   B 18   table cell        1      3       1        1
#> 1.10                  B 18   table cell        1      4       1        1
#> 1.13                  C 18   table cell        1      5       1        1
#> 2.2            Header 2 18   table cell        1      1       2        1
#> 2.5               12.23 18   table cell        1      2       2        1
#> 2.8                1.23 18   table cell        1      3       2        1
#> 2.11                9.0 18   table cell        1      4       2        1
#> 2.14                  6 18   table cell        1      5       2        1
#> 3.3            Header 3 18   table cell        1      1       3        1
#> 3.6           blah blah 18   table cell        1      2       3        1
#> 3.9      blah blah blah 18   table cell        1      3       3        1
#> 3.12              Salut 18   table cell        1      4       3        1
#> 3.15              Hello 18   table cell        1      5       3        1
#> 11               R logo 15    paragraph        1     NA      NA       NA
#> 12                 <NA> 17        image        1     NA      NA       NA
#> 13                   Hi  2    paragraph        2     NA      NA       NA
#> 21             This is   3    paragraph        2     NA      NA       NA
#> 31         an unordered  3    paragraph        2     NA      NA       NA
#> 41   list of paragraphs  3    paragraph        2     NA      NA       NA
#> 51                       3    paragraph        2     NA      NA       NA
#> 6    This is an ordered  3    paragraph        2     NA      NA       NA
#> 7    list of paragraphs  3    paragraph        2     NA      NA       NA
#> 1.12                     5   table cell        3      1       1        1
#> 1.6                   a  5   table cell        3      2       1        2
#> 1.11                     5   table cell        3      3       1        1
#> 1.16                  b  5   table cell        3      4       1        1
#> 1.21               <NA>  5   table cell        3      5       1        1
#> 2.21                     5   table cell        3      1       2        1
#> 2.7                <NA>  5   table cell        3      2       2        0
#> 2.12                     5   table cell        3      3       2        1
#> 2.17                     5   table cell        3      4       2        1
#> 2.22                     5   table cell        3      5       2        1
#> 3.31                     5   table cell        3      1       3        1
#> 3.8                      5   table cell        3      2       3        1
#> 3.13                  c  5   table cell        3      3       3        1
#> 3.18               <NA>  5   table cell        3      4       3        1
#> 3.23                     5   table cell        3      5       3        1
#> 4.4                      5   table cell        3      1       4        1
#> 4.9                      5   table cell        3      2       4        1
#> 4.14                     5   table cell        3      3       4        1
#> 4.19                     5   table cell        3      4       4        1
#> 4.24                     5   table cell        3      5       4        1
#> 5.5                      5   table cell        3      1       5        1
#> 5.10                  d  5   table cell        3      2       5        1
#> 5.15               <NA>  5   table cell        3      3       5        1
#> 5.20                     5   table cell        3      4       5        1
#> 5.25                     5   table cell        3      5       5        1
#>      row_span            media_file
#> 1          NA                  <NA>
#> 2          NA                  <NA>
#> 3          NA                  <NA>
#> 4          NA                  <NA>
#> 5          NA                  <NA>
#> 1.1         1                  <NA>
#> 1.4         1                  <NA>
#> 1.7         1                  <NA>
#> 1.10        1                  <NA>
#> 1.13        1                  <NA>
#> 2.2         1                  <NA>
#> 2.5         1                  <NA>
#> 2.8         1                  <NA>
#> 2.11        1                  <NA>
#> 2.14        1                  <NA>
#> 3.3         1                  <NA>
#> 3.6         1                  <NA>
#> 3.9         1                  <NA>
#> 3.12        1                  <NA>
#> 3.15        1                  <NA>
#> 11         NA                  <NA>
#> 12         NA ppt/media//image1.png
#> 13         NA                  <NA>
#> 21         NA                  <NA>
#> 31         NA                  <NA>
#> 41         NA                  <NA>
#> 51         NA                  <NA>
#> 6          NA                  <NA>
#> 7          NA                  <NA>
#> 1.12        1                  <NA>
#> 1.6         1                  <NA>
#> 1.11        1                  <NA>
#> 1.16        2                  <NA>
#> 1.21        0                  <NA>
#> 2.21        1                  <NA>
#> 2.7         1                  <NA>
#> 2.12        1                  <NA>
#> 2.17        1                  <NA>
#> 2.22        1                  <NA>
#> 3.31        1                  <NA>
#> 3.8         1                  <NA>
#> 3.13        2                  <NA>
#> 3.18        0                  <NA>
#> 3.23        1                  <NA>
#> 4.4         1                  <NA>
#> 4.9         1                  <NA>
#> 4.14        1                  <NA>
#> 4.19        1                  <NA>
#> 4.24        1                  <NA>
#> 5.5         1                  <NA>
#> 5.10        2                  <NA>
#> 5.15        0                  <NA>
#> 5.20        1                  <NA>
#> 5.25        1                  <NA>
```
