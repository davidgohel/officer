# Get Word content in a data.frame

read content of a Word document and return a data.frame representing the
document.

## Usage

``` r
docx_summary(x, preserve = FALSE, remove_fields = FALSE, detailed = FALSE)
```

## Arguments

- x:

  an rdocx object

- preserve:

  If `FALSE` (default), text in table cells is collapsed into a single
  line. If `TRUE`, line breaks in table cells are preserved as a "\n"
  character. This feature is adapted from
  `docxtractr::docx_extract_tbl()` published under a [MIT
  licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE)
  in the 'docxtractr' package by Bob Rudis.

- remove_fields:

  if TRUE, prevent field codes from appearing in the returned
  data.frame.

- detailed:

  Should run-level information be included in the dataframe? Defaults to
  `FALSE`. If `TRUE`, the dataframe contains detailed information about
  each run (text formatting, images, hyperlinks, etc.) instead of
  collapsing content at the paragraph level. When `FALSE`, run-level
  information such as images, hyperlinks, and text formatting is not
  available since data is aggregated at the paragraph level.

## Value

A data.frame with the following columns depending on the value of
`detailed`:

When `detailed = FALSE` (default), the data.frame contains:

- `doc_index`: Document element index (integer).

- `content_type`: Type of content: "paragraph" or "table cell"
  (character).

- `style_name`: Name of the paragraph style (character).

- `text`: Collapsed text content of the paragraph or cell (character).

- `table_index`: Index of the table (integer). `NA` for non-table
  content.

- `row_id`: Row position in table (integer). `NA` for non-table content.

- `cell_id`: Cell position in table row (integer). `NA` for non-table
  content.

- `is_header`: Whether the row is a table header (logical). `NA` for
  non-table content.

- `row_span`: Number of rows spanned by the cell (integer). `0` for
  merged cells. `NA` for non-table content.

- `col_span`: Number of columns spanned by the cell (character). `NA`
  for non-table content.

- `table_stylename`: Name of the table style (character). `NA` for
  non-table content.

When `detailed = TRUE`, the data.frame contains additional run-level
information:

- `run_index`: Index of the run within the paragraph (integer).

- `run_content_index`: Index of content element within the run
  (integer).

- `run_content_text`: Text content of the run element (character).

- `image_path`: Path to embedded image stored in the temporary directory
  associated with the rdocx object (character). Images should be copied
  to a permanent location before closing the R session if needed.

- `field_code`: Field code content (character).

- `footnote_text`: Footnote text content (character).

- `link`: Hyperlink URL (character).

- `link_to_bookmark`: Internal bookmark anchor name for hyperlinks
  (character).

- `bookmark_start`: Names of the bookmarks starting on this paragraph
  (values are concatenated with '\|').

- `character_stylename`: Name of the character/run style (character).

- `sz`: Font size in half-points (integer).

- `sz_cs`: Complex script font size in half-points (integer).

- `font_family_ascii`: Font family for ASCII characters (character).

- `font_family_eastasia`: Font family for East Asian characters
  (character).

- `font_family_hansi`: Font family for high ANSI characters (character).

- `font_family_cs`: Font family for complex script characters
  (character).

- `bold`: Whether the run is bold (logical).

- `italic`: Whether the run is italic (logical).

- `underline`: Whether the run is underlined (logical).

- `color`: Text color in hexadecimal format (character).

- `shading`: Shading pattern (character).

- `shading_color`: Shading foreground color (character).

- `shading_fill`: Shading background fill color (character).

- `keep_with_next`: Whether paragraph should stay with next (logical).

- `align`: Paragraph alignment (character).

- `level`: Numbering level (integer). `NA` if not a numbered list.

- `num_id`: Numbering definition ID (integer). `NA` if not a numbered
  list.

## Note

Documents included with
[`body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.md)
will not be accessible in the results.

## Examples

``` r
library(officer)

example_docx <- system.file(
  package = "officer",
  "doc_examples/example.docx"
)
doc <- read_docx(example_docx)

docx_summary(doc)
#>    doc_index content_type     style_name
#> 1          1    paragraph      heading 1
#> 2          2    paragraph           <NA>
#> 3          3    paragraph      heading 1
#> 4          4    paragraph List Paragraph
#> 5          5    paragraph List Paragraph
#> 6          6    paragraph List Paragraph
#> 7          7    paragraph      heading 2
#> 8          8    paragraph List Paragraph
#> 9          9    paragraph List Paragraph
#> 10        10    paragraph List Paragraph
#> 11        12    paragraph           <NA>
#> 12        13    paragraph      heading 2
#> 13        14    paragraph           <NA>
#> 14        16   table cell           <NA>
#> 15        17   table cell           <NA>
#> 16        18   table cell           <NA>
#> 17        19   table cell           <NA>
#> 18        20   table cell           <NA>
#> 19        21   table cell           <NA>
#> 20        22   table cell           <NA>
#> 21        23   table cell           <NA>
#> 22        24   table cell           <NA>
#> 23        25   table cell           <NA>
#> 24        26   table cell           <NA>
#> 25        27   table cell           <NA>
#> 26        29   table cell           <NA>
#> 27        30   table cell           <NA>
#> 28        31   table cell           <NA>
#> 29        32   table cell           <NA>
#> 30        33   table cell           <NA>
#> 31        34   table cell           <NA>
#> 32        35   table cell           <NA>
#> 33        36   table cell           <NA>
#> 34        37   table cell           <NA>
#> 35        38   table cell           <NA>
#> 36        39   table cell           <NA>
#> 37        40   table cell           <NA>
#> 38        41   table cell           <NA>
#> 39        43   table cell           <NA>
#> 40        44   table cell           <NA>
#> 41        45   table cell           <NA>
#> 42        47   table cell           <NA>
#> 43        49   table cell           <NA>
#> 44        50   table cell           <NA>
#> 45        51   table cell           <NA>
#> 46        53   table cell           <NA>
#> 47        54   table cell           <NA>
#> 48        55   table cell           <NA>
#> 49        57   table cell           <NA>
#> 50        58   table cell           <NA>
#> 51        59   table cell           <NA>
#> 52        60   table cell           <NA>
#> 53        61   table cell           <NA>
#> 54        62   table cell           <NA>
#> 55        63   table cell           <NA>
#> 56        64   table cell           <NA>
#>                                                                       text
#> 1                                                                  Title 1
#> 2                Lorem ipsum dolor sit amet, consectetur adipiscing elit. 
#> 3                                                                  Title 2
#> 4                                                       Quisque tristique 
#> 5                                                Augue nisi, et convallis 
#> 6                                                      Sapien mollis nec. 
#> 7                                                              Sub title 1
#> 8                                                       Quisque tristique 
#> 9                                                Augue nisi, et convallis 
#> 10                                                     Sapien mollis nec. 
#> 11          Phasellus nec nunc vitae nulla interdum volutpat eu ac massa. 
#> 12                                                             Sub title 2
#> 13 Morbi rhoncus sapien sit amet leo eleifend, vel fermentum nisi mattis. 
#> 14                                                                  Petals
#> 15                                                               Internode
#> 16                                                                   Sepal
#> 17                                                                   Bract
#> 18                                                             5,621498349
#> 19                                                             2,462106579
#> 20                                                              18,2034091
#> 21                                                             4,994616997
#> 22                                                                      AA
#> 23                                                             2,429320759
#> 24                                                             17,65204912
#> 25                                                             4,767504884
#> 26                                                                     AAA
#> 27                                                              25,9242382
#> 28                                                             2,066051345
#> 29                                                             18,37915478
#> 30                                                             6,489375001
#> 31                                                             25,21130805
#> 32                                                             2,901582763
#> 33                                                             17,31304737
#> 34                                                             17,07215724
#> 35                                                              18,2902189
#> 36                                                               5,7858682
#> 37                                                             25,52433147
#> 38                                                             2,655642742
#> 39                                                             5,645575295
#> 40                                                             Merged cell
#> 41                                                             2,278691288
#> 42                                                             4,828953215
#> 43                                                             2,238467716
#> 44                                                             19,87376227
#> 45                                                             6,783500773
#> 46                                                             2,202762147
#> 47                                                             19,85326662
#> 48                                                             5,395076839
#> 49                                                             2,538375992
#> 50                                                             19,56545356
#> 51                                                             4,683617783
#> 52                                                              29,2459239
#> 53                                                             2,601945544
#> 54                                                             18,95335451
#> 55                                                                    Note
#> 56                                                           New line note
#>    table_index row_id cell_id is_header row_span col_span table_stylename
#> 1           NA     NA      NA        NA       NA     <NA>            <NA>
#> 2           NA     NA      NA        NA       NA     <NA>            <NA>
#> 3           NA     NA      NA        NA       NA     <NA>            <NA>
#> 4           NA     NA      NA        NA       NA     <NA>            <NA>
#> 5           NA     NA      NA        NA       NA     <NA>            <NA>
#> 6           NA     NA      NA        NA       NA     <NA>            <NA>
#> 7           NA     NA      NA        NA       NA     <NA>            <NA>
#> 8           NA     NA      NA        NA       NA     <NA>            <NA>
#> 9           NA     NA      NA        NA       NA     <NA>            <NA>
#> 10          NA     NA      NA        NA       NA     <NA>            <NA>
#> 11          NA     NA      NA        NA       NA     <NA>            <NA>
#> 12          NA     NA      NA        NA       NA     <NA>            <NA>
#> 13          NA     NA      NA        NA       NA     <NA>            <NA>
#> 14           1      1       1      TRUE        1        1   Light Shading
#> 15           1      1       2      TRUE        1        1   Light Shading
#> 16           1      1       3      TRUE        1        1   Light Shading
#> 17           1      1       4      TRUE        1        1   Light Shading
#> 18           1      2       1     FALSE        1        2   Light Shading
#> 19           1      2       3     FALSE        1        2   Light Shading
#> 20           1      2       4     FALSE        1        2   Light Shading
#> 21           1      3       1     FALSE        1        1   Light Shading
#> 22           1      3       2     FALSE        2        1   Light Shading
#> 23           1      3       3     FALSE        1        1   Light Shading
#> 24           1      3       4     FALSE        1        1   Light Shading
#> 25           1      4       1     FALSE        1        1   Light Shading
#> 26           1      4       3     FALSE        1        2   Light Shading
#> 27           1      5       1     FALSE        1        2   Light Shading
#> 28           1      5       3     FALSE        1        1   Light Shading
#> 29           1      5       4     FALSE        1        1   Light Shading
#> 30           1      6       1     FALSE        1        1   Light Shading
#> 31           1      6       2     FALSE        1        1   Light Shading
#> 32           1      6       3     FALSE        1        1   Light Shading
#> 33           1      6       4     FALSE        1        1   Light Shading
#> 34           1      6       4     FALSE        1        1   Light Shading
#> 35           1      6       4     FALSE        3        1   Light Shading
#> 36           1      7       1     FALSE        1        1   Light Shading
#> 37           1      7       2     FALSE        1        1   Light Shading
#> 38           1      7       3     FALSE        1        1   Light Shading
#> 39           1      8       1     FALSE        1        1   Light Shading
#> 40           1      8       2     FALSE        4        1   Light Shading
#> 41           1      8       3     FALSE        1        1   Light Shading
#> 42           1      9       1     FALSE        1        1   Light Shading
#> 43           1      9       3     FALSE        1        1   Light Shading
#> 44           1      9       4     FALSE        1        1   Light Shading
#> 45           1     10       1     FALSE        1        1   Light Shading
#> 46           1     10       3     FALSE        1        1   Light Shading
#> 47           1     10       4     FALSE        1        1   Light Shading
#> 48           1     11       1     FALSE        1        1   Light Shading
#> 49           1     11       3     FALSE        1        1   Light Shading
#> 50           1     11       4     FALSE        1        1   Light Shading
#> 51           1     12       1     FALSE        1        1   Light Shading
#> 52           1     12       2     FALSE        1        1   Light Shading
#> 53           1     12       3     FALSE        1        1   Light Shading
#> 54           1     12       4     FALSE        1        1   Light Shading
#> 55           1     13       1     FALSE        1        4   Light Shading
#> 56           1     13       4     FALSE        1        4   Light Shading

docx_summary(doc, detailed = TRUE)
#>     doc_index content_type run_index run_content_index    run_content_text
#> 1           1    paragraph         1                 1               Title
#> 2           1    paragraph         2                 1                   1
#> 3           2    paragraph         1                 1        Lorem ipsum 
#> 4           2    paragraph         2                 1               dolor
#> 5           2    paragraph         3                 1                    
#> 6           2    paragraph         4                 1                 sit
#> 7           2    paragraph         5                 1                    
#> 8           2    paragraph         6                 1                amet
#> 9           2    paragraph         7                 1                  , 
#> 10          2    paragraph         8                 1         consectetur
#> 11          2    paragraph         9                 1                    
#> 12          2    paragraph        10                 1          adipiscing
#> 13          2    paragraph        11                 1                    
#> 14          2    paragraph        12                 1                elit
#> 15          2    paragraph        13                 1                  . 
#> 16          3    paragraph         1                 1               Title
#> 17          3    paragraph         2                 1                   2
#> 18          4    paragraph         1                 1             Quisque
#> 19          4    paragraph         2                 1          tristique 
#> 20          5    paragraph         1                 1                   A
#> 21          5    paragraph         2                 1                ugue
#> 22          5    paragraph         3                 1                    
#> 23          5    paragraph         4                 1                nisi
#> 24          5    paragraph         5                 1               , et 
#> 25          5    paragraph         6                 1           convallis
#> 26          5    paragraph         7                 1                    
#> 27          6    paragraph         1                 1                   S
#> 28          6    paragraph         2                 1  apien mollis nec. 
#> 29          7    paragraph         1                 1                 Sub
#> 30          7    paragraph         2                 1                    
#> 31          7    paragraph         3                 1               title
#> 32          7    paragraph         4                 1                   1
#> 33          8    paragraph         1                 1             Quisque
#> 34          8    paragraph         2                 1          tristique 
#> 35          9    paragraph         1                 1               Augue
#> 36          9    paragraph         2                 1                    
#> 37          9    paragraph         3                 1                nisi
#> 38          9    paragraph         4                 1               , et 
#> 39          9    paragraph         5                 1           convallis
#> 40          9    paragraph         6                 1                    
#> 41         10    paragraph         1                 1 Sapien mollis nec. 
#> 42         12    paragraph         1                 1           Phasellus
#> 43         12    paragraph         2                 1     nec nunc vitae 
#> 44         12    paragraph         3                 1               nulla
#> 45         12    paragraph         4                 1                    
#> 46         12    paragraph         5                 1            interdum
#> 47         12    paragraph         6                 1                    
#> 48         12    paragraph         7                 1            volutpat
#> 49         12    paragraph         8                 1                 eu 
#> 50         12    paragraph         9                 1                  ac
#> 51         12    paragraph        10                 1             massa. 
#> 52         13    paragraph         1                 1                 Sub
#> 53         13    paragraph         2                 1                    
#> 54         13    paragraph         3                 1               title
#> 55         13    paragraph         4                 1                   2
#> 56         14    paragraph         1                 1      Morbi rhoncus 
#> 57         14    paragraph         2                 1              sapien
#> 58         14    paragraph         3                 1                    
#> 59         14    paragraph         4                 1                 sit
#> 60         14    paragraph         5                 1                    
#> 61         14    paragraph         6                 1                amet
#> 62         14    paragraph         7                 1                    
#> 63         14    paragraph         8                 1                 leo
#> 64         14    paragraph         9                 1                    
#> 65         14    paragraph        10                 1            eleifend
#> 66         14    paragraph        11                 1                  , 
#> 67         14    paragraph        12                 1                 vel
#> 68         14    paragraph        13                 1                    
#> 69         14    paragraph        14                 1           fermentum
#> 70         14    paragraph        15                 1                    
#> 71         14    paragraph        16                 1                nisi
#> 72         14    paragraph        17                 1                    
#> 73         14    paragraph        18                 1              mattis
#> 74         14    paragraph        19                 1                  . 
#> 75         16   table cell         1                 1              Petals
#> 76         17   table cell         1                 1           Internode
#> 77         18   table cell         1                 1               Sepal
#> 78         19   table cell         1                 1               Bract
#> 79         20   table cell         1                 1         5,621498349
#> 80         21   table cell         1                 1         2,462106579
#> 81         22   table cell         1                 1          18,2034091
#> 82         23   table cell         1                 1         4,994616997
#> 83         24   table cell         1                 1                  AA
#> 84         25   table cell         1                 1         2,429320759
#> 85         26   table cell         1                 1         17,65204912
#> 86         27   table cell         1                 1         4,767504884
#> 87         29   table cell         1                 1                 AAA
#> 88         30   table cell         1                 1          25,9242382
#> 89         31   table cell         1                 1         2,066051345
#> 90         32   table cell         1                 1         18,37915478
#> 91         33   table cell         1                 1         6,489375001
#> 92         34   table cell         1                 1         25,21130805
#> 93         35   table cell         1                 1         2,901582763
#> 94         36   table cell         1                 1         17,31304737
#> 95         37   table cell         1                 1         17,07215724
#> 96         38   table cell         1                 1          18,2902189
#> 97         39   table cell         1                 1           5,7858682
#> 98         40   table cell         1                 1         25,52433147
#> 99         41   table cell         1                 1         2,655642742
#> 100        43   table cell         1                 1         5,645575295
#> 101        44   table cell         1                 1              Merged
#> 102        44   table cell         2                 1                    
#> 103        44   table cell         3                 1                cell
#> 104        45   table cell         1                 1         2,278691288
#> 105        47   table cell         1                 1         4,828953215
#> 106        49   table cell         1                 1         2,238467716
#> 107        50   table cell         1                 1         19,87376227
#> 108        51   table cell         1                 1         6,783500773
#> 109        53   table cell         1                 1         2,202762147
#> 110        54   table cell         1                 1         19,85326662
#> 111        55   table cell         1                 1         5,395076839
#> 112        57   table cell         1                 1         2,538375992
#> 113        58   table cell         1                 1         19,56545356
#> 114        59   table cell         1                 1         4,683617783
#> 115        60   table cell         1                 1          29,2459239
#> 116        61   table cell         1                 1         2,601945544
#> 117        62   table cell         1                 1         18,95335451
#> 118        63   table cell         1                 1                Note
#> 119        64   table cell         1                 1       New line note
#>     image_path field_code footnote_text link link_to_bookmark bookmark_start
#> 1         <NA>       <NA>               <NA>             <NA>           <NA>
#> 2         <NA>       <NA>               <NA>             <NA>           <NA>
#> 3         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 4         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 5         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 6         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 7         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 8         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 9         <NA>       <NA>               <NA>             <NA>          bmk_1
#> 10        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 11        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 12        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 13        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 14        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 15        <NA>       <NA>               <NA>             <NA>          bmk_1
#> 16        <NA>       <NA>               <NA>             <NA>           <NA>
#> 17        <NA>       <NA>               <NA>             <NA>           <NA>
#> 18        <NA>       <NA>               <NA>             <NA>           <NA>
#> 19        <NA>       <NA>               <NA>             <NA>           <NA>
#> 20        <NA>       <NA>               <NA>             <NA>           <NA>
#> 21        <NA>       <NA>               <NA>             <NA>           <NA>
#> 22        <NA>       <NA>               <NA>             <NA>           <NA>
#> 23        <NA>       <NA>               <NA>             <NA>           <NA>
#> 24        <NA>       <NA>               <NA>             <NA>           <NA>
#> 25        <NA>       <NA>               <NA>             <NA>           <NA>
#> 26        <NA>       <NA>               <NA>             <NA>           <NA>
#> 27        <NA>       <NA>               <NA>             <NA>           <NA>
#> 28        <NA>       <NA>               <NA>             <NA>           <NA>
#> 29        <NA>       <NA>               <NA>             <NA>           <NA>
#> 30        <NA>       <NA>               <NA>             <NA>           <NA>
#> 31        <NA>       <NA>               <NA>             <NA>           <NA>
#> 32        <NA>       <NA>               <NA>             <NA>           <NA>
#> 33        <NA>       <NA>               <NA>             <NA>           <NA>
#> 34        <NA>       <NA>               <NA>             <NA>           <NA>
#> 35        <NA>       <NA>               <NA>             <NA>           <NA>
#> 36        <NA>       <NA>               <NA>             <NA>           <NA>
#> 37        <NA>       <NA>               <NA>             <NA>           <NA>
#> 38        <NA>       <NA>               <NA>             <NA>           <NA>
#> 39        <NA>       <NA>               <NA>             <NA>           <NA>
#> 40        <NA>       <NA>               <NA>             <NA>           <NA>
#> 41        <NA>       <NA>               <NA>             <NA>           <NA>
#> 42        <NA>       <NA>               <NA>             <NA>           <NA>
#> 43        <NA>       <NA>               <NA>             <NA>           <NA>
#> 44        <NA>       <NA>               <NA>             <NA>           <NA>
#> 45        <NA>       <NA>               <NA>             <NA>           <NA>
#> 46        <NA>       <NA>               <NA>             <NA>           <NA>
#> 47        <NA>       <NA>               <NA>             <NA>           <NA>
#> 48        <NA>       <NA>               <NA>             <NA>           <NA>
#> 49        <NA>       <NA>               <NA>             <NA>           <NA>
#> 50        <NA>       <NA>               <NA>             <NA>           <NA>
#> 51        <NA>       <NA>               <NA>             <NA>           <NA>
#> 52        <NA>       <NA>               <NA>             <NA>           <NA>
#> 53        <NA>       <NA>               <NA>             <NA>           <NA>
#> 54        <NA>       <NA>               <NA>             <NA>           <NA>
#> 55        <NA>       <NA>               <NA>             <NA>           <NA>
#> 56        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 57        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 58        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 59        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 60        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 61        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 62        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 63        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 64        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 65        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 66        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 67        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 68        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 69        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 70        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 71        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 72        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 73        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 74        <NA>       <NA>               <NA>             <NA>          bmk_2
#> 75        <NA>       <NA>               <NA>             <NA>           <NA>
#> 76        <NA>       <NA>               <NA>             <NA>           <NA>
#> 77        <NA>       <NA>               <NA>             <NA>           <NA>
#> 78        <NA>       <NA>               <NA>             <NA>           <NA>
#> 79        <NA>       <NA>               <NA>             <NA>           <NA>
#> 80        <NA>       <NA>               <NA>             <NA>           <NA>
#> 81        <NA>       <NA>               <NA>             <NA>           <NA>
#> 82        <NA>       <NA>               <NA>             <NA>           <NA>
#> 83        <NA>       <NA>               <NA>             <NA>           <NA>
#> 84        <NA>       <NA>               <NA>             <NA>           <NA>
#> 85        <NA>       <NA>               <NA>             <NA>           <NA>
#> 86        <NA>       <NA>               <NA>             <NA>           <NA>
#> 87        <NA>       <NA>               <NA>             <NA>           <NA>
#> 88        <NA>       <NA>               <NA>             <NA>           <NA>
#> 89        <NA>       <NA>               <NA>             <NA>           <NA>
#> 90        <NA>       <NA>               <NA>             <NA>           <NA>
#> 91        <NA>       <NA>               <NA>             <NA>           <NA>
#> 92        <NA>       <NA>               <NA>             <NA>           <NA>
#> 93        <NA>       <NA>               <NA>             <NA>           <NA>
#> 94        <NA>       <NA>               <NA>             <NA>           <NA>
#> 95        <NA>       <NA>               <NA>             <NA>           <NA>
#> 96        <NA>       <NA>               <NA>             <NA>           <NA>
#> 97        <NA>       <NA>               <NA>             <NA>           <NA>
#> 98        <NA>       <NA>               <NA>             <NA>           <NA>
#> 99        <NA>       <NA>               <NA>             <NA>           <NA>
#> 100       <NA>       <NA>               <NA>             <NA>           <NA>
#> 101       <NA>       <NA>               <NA>             <NA>           <NA>
#> 102       <NA>       <NA>               <NA>             <NA>           <NA>
#> 103       <NA>       <NA>               <NA>             <NA>           <NA>
#> 104       <NA>       <NA>               <NA>             <NA>           <NA>
#> 105       <NA>       <NA>               <NA>             <NA>           <NA>
#> 106       <NA>       <NA>               <NA>             <NA>           <NA>
#> 107       <NA>       <NA>               <NA>             <NA>           <NA>
#> 108       <NA>       <NA>               <NA>             <NA>           <NA>
#> 109       <NA>       <NA>               <NA>             <NA>           <NA>
#> 110       <NA>       <NA>               <NA>             <NA>           <NA>
#> 111       <NA>       <NA>               <NA>             <NA>           <NA>
#> 112       <NA>       <NA>               <NA>             <NA>           <NA>
#> 113       <NA>       <NA>               <NA>             <NA>           <NA>
#> 114       <NA>       <NA>               <NA>             <NA>           <NA>
#> 115       <NA>       <NA>               <NA>             <NA>           <NA>
#> 116       <NA>       <NA>               <NA>             <NA>           <NA>
#> 117       <NA>       <NA>               <NA>             <NA>           <NA>
#> 118       <NA>       <NA>               <NA>             <NA>           <NA>
#> 119       <NA>       <NA>               <NA>             <NA>           <NA>
#>     character_stylename sz sz_cs font_family_ascii font_family_eastasia
#> 1                  <NA> NA    NA              <NA>                 <NA>
#> 2                  <NA> NA    NA              <NA>                 <NA>
#> 3                  <NA> NA    NA              <NA>                 <NA>
#> 4                  <NA> NA    NA              <NA>                 <NA>
#> 5                  <NA> NA    NA              <NA>                 <NA>
#> 6                  <NA> NA    NA              <NA>                 <NA>
#> 7                  <NA> NA    NA              <NA>                 <NA>
#> 8                  <NA> NA    NA              <NA>                 <NA>
#> 9                  <NA> NA    NA              <NA>                 <NA>
#> 10                 <NA> NA    NA              <NA>                 <NA>
#> 11                 <NA> NA    NA              <NA>                 <NA>
#> 12                 <NA> NA    NA              <NA>                 <NA>
#> 13                 <NA> NA    NA              <NA>                 <NA>
#> 14                 <NA> NA    NA              <NA>                 <NA>
#> 15                 <NA> NA    NA              <NA>                 <NA>
#> 16                 <NA> NA    NA              <NA>                 <NA>
#> 17                 <NA> NA    NA              <NA>                 <NA>
#> 18                 <NA> NA    NA              <NA>                 <NA>
#> 19                 <NA> NA    NA              <NA>                 <NA>
#> 20                 <NA> NA    NA              <NA>                 <NA>
#> 21                 <NA> NA    NA              <NA>                 <NA>
#> 22                 <NA> NA    NA              <NA>                 <NA>
#> 23                 <NA> NA    NA              <NA>                 <NA>
#> 24                 <NA> NA    NA              <NA>                 <NA>
#> 25                 <NA> NA    NA              <NA>                 <NA>
#> 26                 <NA> NA    NA              <NA>                 <NA>
#> 27                 <NA> NA    NA              <NA>                 <NA>
#> 28                 <NA> NA    NA              <NA>                 <NA>
#> 29                 <NA> NA    NA              <NA>                 <NA>
#> 30                 <NA> NA    NA              <NA>                 <NA>
#> 31                 <NA> NA    NA              <NA>                 <NA>
#> 32                 <NA> NA    NA              <NA>                 <NA>
#> 33                 <NA> NA    NA              <NA>                 <NA>
#> 34                 <NA> NA    NA              <NA>                 <NA>
#> 35                 <NA> NA    NA              <NA>                 <NA>
#> 36                 <NA> NA    NA              <NA>                 <NA>
#> 37                 <NA> NA    NA              <NA>                 <NA>
#> 38                 <NA> NA    NA              <NA>                 <NA>
#> 39                 <NA> NA    NA              <NA>                 <NA>
#> 40                 <NA> NA    NA              <NA>                 <NA>
#> 41                 <NA> NA    NA              <NA>                 <NA>
#> 42                 <NA> NA    NA              <NA>                 <NA>
#> 43                 <NA> NA    NA              <NA>                 <NA>
#> 44                 <NA> NA    NA              <NA>                 <NA>
#> 45                 <NA> NA    NA              <NA>                 <NA>
#> 46                 <NA> NA    NA              <NA>                 <NA>
#> 47                 <NA> NA    NA              <NA>                 <NA>
#> 48                 <NA> NA    NA              <NA>                 <NA>
#> 49                 <NA> NA    NA              <NA>                 <NA>
#> 50                 <NA> NA    NA              <NA>                 <NA>
#> 51                 <NA> NA    NA              <NA>                 <NA>
#> 52                 <NA> NA    NA              <NA>                 <NA>
#> 53                 <NA> NA    NA              <NA>                 <NA>
#> 54                 <NA> NA    NA              <NA>                 <NA>
#> 55                 <NA> NA    NA              <NA>                 <NA>
#> 56                 <NA> NA    NA              <NA>                 <NA>
#> 57                 <NA> NA    NA              <NA>                 <NA>
#> 58                 <NA> NA    NA              <NA>                 <NA>
#> 59                 <NA> NA    NA              <NA>                 <NA>
#> 60                 <NA> NA    NA              <NA>                 <NA>
#> 61                 <NA> NA    NA              <NA>                 <NA>
#> 62                 <NA> NA    NA              <NA>                 <NA>
#> 63                 <NA> NA    NA              <NA>                 <NA>
#> 64                 <NA> NA    NA              <NA>                 <NA>
#> 65                 <NA> NA    NA              <NA>                 <NA>
#> 66                 <NA> NA    NA              <NA>                 <NA>
#> 67                 <NA> NA    NA              <NA>                 <NA>
#> 68                 <NA> NA    NA              <NA>                 <NA>
#> 69                 <NA> NA    NA              <NA>                 <NA>
#> 70                 <NA> NA    NA              <NA>                 <NA>
#> 71                 <NA> NA    NA              <NA>                 <NA>
#> 72                 <NA> NA    NA              <NA>                 <NA>
#> 73                 <NA> NA    NA              <NA>                 <NA>
#> 74                 <NA> NA    NA              <NA>                 <NA>
#> 75                 <NA> NA    NA           Calibri      Times New Roman
#> 76                 <NA> NA    NA           Calibri      Times New Roman
#> 77                 <NA> NA    NA           Calibri      Times New Roman
#> 78                 <NA> NA    NA           Calibri      Times New Roman
#> 79                 <NA> NA    NA           Calibri      Times New Roman
#> 80                 <NA> NA    NA           Calibri      Times New Roman
#> 81                 <NA> NA    NA           Calibri      Times New Roman
#> 82                 <NA> NA    NA           Calibri      Times New Roman
#> 83                 <NA> NA    NA           Calibri      Times New Roman
#> 84                 <NA> NA    NA           Calibri      Times New Roman
#> 85                 <NA> NA    NA           Calibri      Times New Roman
#> 86                 <NA> NA    NA           Calibri      Times New Roman
#> 87                 <NA> NA    NA           Calibri      Times New Roman
#> 88                 <NA> NA    NA           Calibri      Times New Roman
#> 89                 <NA> NA    NA           Calibri      Times New Roman
#> 90                 <NA> NA    NA           Calibri      Times New Roman
#> 91                 <NA> NA    NA           Calibri      Times New Roman
#> 92                 <NA> NA    NA           Calibri      Times New Roman
#> 93                 <NA> NA    NA           Calibri      Times New Roman
#> 94                 <NA> NA    NA           Calibri      Times New Roman
#> 95                 <NA> NA    NA           Calibri      Times New Roman
#> 96                 <NA> NA    NA           Calibri      Times New Roman
#> 97                 <NA> NA    NA           Calibri      Times New Roman
#> 98                 <NA> NA    NA           Calibri      Times New Roman
#> 99                 <NA> NA    NA           Calibri      Times New Roman
#> 100                <NA> NA    NA           Calibri      Times New Roman
#> 101                <NA> NA    NA           Calibri      Times New Roman
#> 102                <NA> NA    NA           Calibri      Times New Roman
#> 103                <NA> NA    NA           Calibri      Times New Roman
#> 104                <NA> NA    NA           Calibri      Times New Roman
#> 105                <NA> NA    NA           Calibri      Times New Roman
#> 106                <NA> NA    NA           Calibri      Times New Roman
#> 107                <NA> NA    NA           Calibri      Times New Roman
#> 108                <NA> NA    NA           Calibri      Times New Roman
#> 109                <NA> NA    NA           Calibri      Times New Roman
#> 110                <NA> NA    NA           Calibri      Times New Roman
#> 111                <NA> NA    NA           Calibri      Times New Roman
#> 112                <NA> NA    NA           Calibri      Times New Roman
#> 113                <NA> NA    NA           Calibri      Times New Roman
#> 114                <NA> NA    NA           Calibri      Times New Roman
#> 115                <NA> NA    NA           Calibri      Times New Roman
#> 116                <NA> NA    NA           Calibri      Times New Roman
#> 117                <NA> NA    NA           Calibri      Times New Roman
#> 118                <NA> NA    NA           Calibri      Times New Roman
#> 119                <NA> NA    NA           Calibri      Times New Roman
#>     font_family_hansi  font_family_cs  bold italic underline   color shading
#> 1                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 2                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 3                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 4                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 5                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 6                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 7                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 8                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 9                <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 10               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 11               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 12               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 13               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 14               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 15               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 16               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 17               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 18               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 19               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 20               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 21               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 22               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 23               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 24               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 25               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 26               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 27               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 28               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 29               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 30               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 31               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 32               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 33               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 34               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 35               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 36               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 37               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 38               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 39               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 40               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 41               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 42               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 43               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 44               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 45               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 46               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 47               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 48               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 49               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 50               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 51               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 52               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 53               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 54               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 55               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 56               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 57               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 58               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 59               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 60               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 61               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 62               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 63               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 64               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 65               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 66               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 67               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 68               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 69               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 70               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 71               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 72               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 73               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 74               <NA>            <NA> FALSE  FALSE     FALSE    <NA>    <NA>
#> 75            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 76            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 77            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 78            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 79            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 80            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 81            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 82            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 83            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 84            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 85            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 86            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 87            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 88            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 89            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 90            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 91            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 92            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 93            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 94            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 95            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 96            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 97            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 98            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 99            Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 100           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 101           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 102           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 103           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 104           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 105           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 106           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 107           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 108           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 109           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 110           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 111           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 112           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 113           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 114           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 115           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 116           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 117           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 118           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#> 119           Calibri Times New Roman FALSE  FALSE     FALSE #000000    <NA>
#>     shading_color shading_fill paragraph_stylename keep_with_next  align level
#> 1            <NA>         <NA>           heading 1          FALSE   <NA>    NA
#> 2            <NA>         <NA>           heading 1          FALSE   <NA>    NA
#> 3            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 4            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 5            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 6            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 7            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 8            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 9            <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 10           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 11           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 12           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 13           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 14           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 15           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 16           <NA>         <NA>           heading 1          FALSE   <NA>    NA
#> 17           <NA>         <NA>           heading 1          FALSE   <NA>    NA
#> 18           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 19           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 20           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 21           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 22           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 23           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 24           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 25           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 26           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 27           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 28           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 29           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 30           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 31           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 32           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 33           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 34           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 35           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 36           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 37           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 38           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 39           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 40           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 41           <NA>         <NA>      List Paragraph          FALSE   <NA>     1
#> 42           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 43           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 44           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 45           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 46           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 47           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 48           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 49           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 50           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 51           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 52           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 53           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 54           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 55           <NA>         <NA>           heading 2          FALSE   <NA>    NA
#> 56           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 57           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 58           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 59           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 60           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 61           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 62           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 63           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 64           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 65           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 66           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 67           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 68           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 69           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 70           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 71           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 72           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 73           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 74           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 75           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 76           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 77           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 78           <NA>         <NA>                <NA>          FALSE   <NA>    NA
#> 79           <NA>         <NA>                <NA>          FALSE  right    NA
#> 80           <NA>         <NA>                <NA>          FALSE  right    NA
#> 81           <NA>         <NA>                <NA>          FALSE  right    NA
#> 82           <NA>         <NA>                <NA>          FALSE  right    NA
#> 83           <NA>         <NA>                <NA>          FALSE  right    NA
#> 84           <NA>         <NA>                <NA>          FALSE  right    NA
#> 85           <NA>         <NA>                <NA>          FALSE  right    NA
#> 86           <NA>         <NA>                <NA>          FALSE  right    NA
#> 87           <NA>         <NA>                <NA>          FALSE center    NA
#> 88           <NA>         <NA>                <NA>          FALSE  right    NA
#> 89           <NA>         <NA>                <NA>          FALSE  right    NA
#> 90           <NA>         <NA>                <NA>          FALSE  right    NA
#> 91           <NA>         <NA>                <NA>          FALSE  right    NA
#> 92           <NA>         <NA>                <NA>          FALSE  right    NA
#> 93           <NA>         <NA>                <NA>          FALSE  right    NA
#> 94           <NA>         <NA>                <NA>          FALSE  right    NA
#> 95           <NA>         <NA>                <NA>          FALSE  right    NA
#> 96           <NA>         <NA>                <NA>          FALSE  right    NA
#> 97           <NA>         <NA>                <NA>          FALSE  right    NA
#> 98           <NA>         <NA>                <NA>          FALSE  right    NA
#> 99           <NA>         <NA>                <NA>          FALSE  right    NA
#> 100          <NA>         <NA>                <NA>          FALSE  right    NA
#> 101          <NA>         <NA>                <NA>          FALSE  right    NA
#> 102          <NA>         <NA>                <NA>          FALSE  right    NA
#> 103          <NA>         <NA>                <NA>          FALSE  right    NA
#> 104          <NA>         <NA>                <NA>          FALSE  right    NA
#> 105          <NA>         <NA>                <NA>          FALSE  right    NA
#> 106          <NA>         <NA>                <NA>          FALSE  right    NA
#> 107          <NA>         <NA>                <NA>          FALSE  right    NA
#> 108          <NA>         <NA>                <NA>          FALSE  right    NA
#> 109          <NA>         <NA>                <NA>          FALSE  right    NA
#> 110          <NA>         <NA>                <NA>          FALSE  right    NA
#> 111          <NA>         <NA>                <NA>          FALSE  right    NA
#> 112          <NA>         <NA>                <NA>          FALSE  right    NA
#> 113          <NA>         <NA>                <NA>          FALSE  right    NA
#> 114          <NA>         <NA>                <NA>          FALSE  right    NA
#> 115          <NA>         <NA>                <NA>          FALSE  right    NA
#> 116          <NA>         <NA>                <NA>          FALSE  right    NA
#> 117          <NA>         <NA>                <NA>          FALSE  right    NA
#> 118          <NA>         <NA>                <NA>          FALSE  right    NA
#> 119          <NA>         <NA>                <NA>          FALSE  right    NA
#>     num_id table_index row_id cell_id col_span row_span is_header
#> 1       NA          NA     NA      NA     <NA>       NA        NA
#> 2       NA          NA     NA      NA     <NA>       NA        NA
#> 3       NA          NA     NA      NA     <NA>       NA        NA
#> 4       NA          NA     NA      NA     <NA>       NA        NA
#> 5       NA          NA     NA      NA     <NA>       NA        NA
#> 6       NA          NA     NA      NA     <NA>       NA        NA
#> 7       NA          NA     NA      NA     <NA>       NA        NA
#> 8       NA          NA     NA      NA     <NA>       NA        NA
#> 9       NA          NA     NA      NA     <NA>       NA        NA
#> 10      NA          NA     NA      NA     <NA>       NA        NA
#> 11      NA          NA     NA      NA     <NA>       NA        NA
#> 12      NA          NA     NA      NA     <NA>       NA        NA
#> 13      NA          NA     NA      NA     <NA>       NA        NA
#> 14      NA          NA     NA      NA     <NA>       NA        NA
#> 15      NA          NA     NA      NA     <NA>       NA        NA
#> 16      NA          NA     NA      NA     <NA>       NA        NA
#> 17      NA          NA     NA      NA     <NA>       NA        NA
#> 18       2          NA     NA      NA     <NA>       NA        NA
#> 19       2          NA     NA      NA     <NA>       NA        NA
#> 20       2          NA     NA      NA     <NA>       NA        NA
#> 21       2          NA     NA      NA     <NA>       NA        NA
#> 22       2          NA     NA      NA     <NA>       NA        NA
#> 23       2          NA     NA      NA     <NA>       NA        NA
#> 24       2          NA     NA      NA     <NA>       NA        NA
#> 25       2          NA     NA      NA     <NA>       NA        NA
#> 26       2          NA     NA      NA     <NA>       NA        NA
#> 27       2          NA     NA      NA     <NA>       NA        NA
#> 28       2          NA     NA      NA     <NA>       NA        NA
#> 29      NA          NA     NA      NA     <NA>       NA        NA
#> 30      NA          NA     NA      NA     <NA>       NA        NA
#> 31      NA          NA     NA      NA     <NA>       NA        NA
#> 32      NA          NA     NA      NA     <NA>       NA        NA
#> 33       1          NA     NA      NA     <NA>       NA        NA
#> 34       1          NA     NA      NA     <NA>       NA        NA
#> 35       1          NA     NA      NA     <NA>       NA        NA
#> 36       1          NA     NA      NA     <NA>       NA        NA
#> 37       1          NA     NA      NA     <NA>       NA        NA
#> 38       1          NA     NA      NA     <NA>       NA        NA
#> 39       1          NA     NA      NA     <NA>       NA        NA
#> 40       1          NA     NA      NA     <NA>       NA        NA
#> 41       1          NA     NA      NA     <NA>       NA        NA
#> 42      NA          NA     NA      NA     <NA>       NA        NA
#> 43      NA          NA     NA      NA     <NA>       NA        NA
#> 44      NA          NA     NA      NA     <NA>       NA        NA
#> 45      NA          NA     NA      NA     <NA>       NA        NA
#> 46      NA          NA     NA      NA     <NA>       NA        NA
#> 47      NA          NA     NA      NA     <NA>       NA        NA
#> 48      NA          NA     NA      NA     <NA>       NA        NA
#> 49      NA          NA     NA      NA     <NA>       NA        NA
#> 50      NA          NA     NA      NA     <NA>       NA        NA
#> 51      NA          NA     NA      NA     <NA>       NA        NA
#> 52      NA          NA     NA      NA     <NA>       NA        NA
#> 53      NA          NA     NA      NA     <NA>       NA        NA
#> 54      NA          NA     NA      NA     <NA>       NA        NA
#> 55      NA          NA     NA      NA     <NA>       NA        NA
#> 56      NA          NA     NA      NA     <NA>       NA        NA
#> 57      NA          NA     NA      NA     <NA>       NA        NA
#> 58      NA          NA     NA      NA     <NA>       NA        NA
#> 59      NA          NA     NA      NA     <NA>       NA        NA
#> 60      NA          NA     NA      NA     <NA>       NA        NA
#> 61      NA          NA     NA      NA     <NA>       NA        NA
#> 62      NA          NA     NA      NA     <NA>       NA        NA
#> 63      NA          NA     NA      NA     <NA>       NA        NA
#> 64      NA          NA     NA      NA     <NA>       NA        NA
#> 65      NA          NA     NA      NA     <NA>       NA        NA
#> 66      NA          NA     NA      NA     <NA>       NA        NA
#> 67      NA          NA     NA      NA     <NA>       NA        NA
#> 68      NA          NA     NA      NA     <NA>       NA        NA
#> 69      NA          NA     NA      NA     <NA>       NA        NA
#> 70      NA          NA     NA      NA     <NA>       NA        NA
#> 71      NA          NA     NA      NA     <NA>       NA        NA
#> 72      NA          NA     NA      NA     <NA>       NA        NA
#> 73      NA          NA     NA      NA     <NA>       NA        NA
#> 74      NA          NA     NA      NA     <NA>       NA        NA
#> 75      NA           1      1       1        1        1      TRUE
#> 76      NA           1      1       2        1        1      TRUE
#> 77      NA           1      1       3        1        1      TRUE
#> 78      NA           1      1       4        1        1      TRUE
#> 79      NA           1      2       1        2        1     FALSE
#> 80      NA           1      2       3        2        1     FALSE
#> 81      NA           1      2       4        2        1     FALSE
#> 82      NA           1      3       1        1        1     FALSE
#> 83      NA           1      3       2        1        2     FALSE
#> 84      NA           1      3       3        1        1     FALSE
#> 85      NA           1      3       4        1        1     FALSE
#> 86      NA           1      4       1        1        1     FALSE
#> 87      NA           1      4       3        2        1     FALSE
#> 88      NA           1      5       1        2        1     FALSE
#> 89      NA           1      5       3        1        1     FALSE
#> 90      NA           1      5       4        1        1     FALSE
#> 91      NA           1      6       1        1        1     FALSE
#> 92      NA           1      6       2        1        1     FALSE
#> 93      NA           1      6       3        1        1     FALSE
#> 94      NA           1      6       4        1        1     FALSE
#> 95      NA           1      6       4        1        1     FALSE
#> 96      NA           1      6       4        1        3     FALSE
#> 97      NA           1      7       1        1        1     FALSE
#> 98      NA           1      7       2        1        1     FALSE
#> 99      NA           1      7       3        1        1     FALSE
#> 100     NA           1      8       1        1        1     FALSE
#> 101     NA           1      8       2        1        4     FALSE
#> 102     NA           1      8       2        1        4     FALSE
#> 103     NA           1      8       2        1        4     FALSE
#> 104     NA           1      8       3        1        1     FALSE
#> 105     NA           1      9       1        1        1     FALSE
#> 106     NA           1      9       3        1        1     FALSE
#> 107     NA           1      9       4        1        1     FALSE
#> 108     NA           1     10       1        1        1     FALSE
#> 109     NA           1     10       3        1        1     FALSE
#> 110     NA           1     10       4        1        1     FALSE
#> 111     NA           1     11       1        1        1     FALSE
#> 112     NA           1     11       3        1        1     FALSE
#> 113     NA           1     11       4        1        1     FALSE
#> 114     NA           1     12       1        1        1     FALSE
#> 115     NA           1     12       2        1        1     FALSE
#> 116     NA           1     12       3        1        1     FALSE
#> 117     NA           1     12       4        1        1     FALSE
#> 118     NA           1     13       1        4        1     FALSE
#> 119     NA           1     13       4        4        1     FALSE
#>     table_stylename
#> 1              <NA>
#> 2              <NA>
#> 3              <NA>
#> 4              <NA>
#> 5              <NA>
#> 6              <NA>
#> 7              <NA>
#> 8              <NA>
#> 9              <NA>
#> 10             <NA>
#> 11             <NA>
#> 12             <NA>
#> 13             <NA>
#> 14             <NA>
#> 15             <NA>
#> 16             <NA>
#> 17             <NA>
#> 18             <NA>
#> 19             <NA>
#> 20             <NA>
#> 21             <NA>
#> 22             <NA>
#> 23             <NA>
#> 24             <NA>
#> 25             <NA>
#> 26             <NA>
#> 27             <NA>
#> 28             <NA>
#> 29             <NA>
#> 30             <NA>
#> 31             <NA>
#> 32             <NA>
#> 33             <NA>
#> 34             <NA>
#> 35             <NA>
#> 36             <NA>
#> 37             <NA>
#> 38             <NA>
#> 39             <NA>
#> 40             <NA>
#> 41             <NA>
#> 42             <NA>
#> 43             <NA>
#> 44             <NA>
#> 45             <NA>
#> 46             <NA>
#> 47             <NA>
#> 48             <NA>
#> 49             <NA>
#> 50             <NA>
#> 51             <NA>
#> 52             <NA>
#> 53             <NA>
#> 54             <NA>
#> 55             <NA>
#> 56             <NA>
#> 57             <NA>
#> 58             <NA>
#> 59             <NA>
#> 60             <NA>
#> 61             <NA>
#> 62             <NA>
#> 63             <NA>
#> 64             <NA>
#> 65             <NA>
#> 66             <NA>
#> 67             <NA>
#> 68             <NA>
#> 69             <NA>
#> 70             <NA>
#> 71             <NA>
#> 72             <NA>
#> 73             <NA>
#> 74             <NA>
#> 75    Light Shading
#> 76    Light Shading
#> 77    Light Shading
#> 78    Light Shading
#> 79    Light Shading
#> 80    Light Shading
#> 81    Light Shading
#> 82    Light Shading
#> 83    Light Shading
#> 84    Light Shading
#> 85    Light Shading
#> 86    Light Shading
#> 87    Light Shading
#> 88    Light Shading
#> 89    Light Shading
#> 90    Light Shading
#> 91    Light Shading
#> 92    Light Shading
#> 93    Light Shading
#> 94    Light Shading
#> 95    Light Shading
#> 96    Light Shading
#> 97    Light Shading
#> 98    Light Shading
#> 99    Light Shading
#> 100   Light Shading
#> 101   Light Shading
#> 102   Light Shading
#> 103   Light Shading
#> 104   Light Shading
#> 105   Light Shading
#> 106   Light Shading
#> 107   Light Shading
#> 108   Light Shading
#> 109   Light Shading
#> 110   Light Shading
#> 111   Light Shading
#> 112   Light Shading
#> 113   Light Shading
#> 114   Light Shading
#> 115   Light Shading
#> 116   Light Shading
#> 117   Light Shading
#> 118   Light Shading
#> 119   Light Shading
```
