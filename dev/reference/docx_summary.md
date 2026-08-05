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

- `para_id`: Unique paragraph id (character), can be used to make a join
  with the `para_id` column of
  [`docx_comments()`](https://davidgohel.github.io/officer/dev/reference/docx_comments.md)
  results. `NA` when the paragraph has no id, e.g. in documents not
  produced by 'Word'.

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
[`body_add_docx()`](https://davidgohel.github.io/officer/dev/reference/body_add_docx.md)
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
#>    doc_index  para_id content_type     style_name
#> 1          1 4248496C    paragraph      heading 1
#> 2          2 6A358887    paragraph           <NA>
#> 3          3 22C38A25    paragraph      heading 1
#> 4          4 0358E421    paragraph List Paragraph
#> 5          5 70C2F11E    paragraph List Paragraph
#> 6          6 429784E2    paragraph List Paragraph
#> 7          7 49D1AAF7    paragraph      heading 2
#> 8          8 2F68EB53    paragraph List Paragraph
#> 9          9 6718A3AB    paragraph List Paragraph
#> 10        10 374BD321    paragraph List Paragraph
#> 11        12 0B98AFE8    paragraph           <NA>
#> 12        13 7BA11984    paragraph      heading 2
#> 13        14 6E5CCCAB    paragraph           <NA>
#> 14        16 105C9948   table cell           <NA>
#> 15        17 4B712A9C   table cell           <NA>
#> 16        18 4F7C7923   table cell           <NA>
#> 17        19 75402514   table cell           <NA>
#> 18        20 0AA8CF0C   table cell           <NA>
#> 19        21 07F46688   table cell           <NA>
#> 20        22 058EDA18   table cell           <NA>
#> 21        23 529A35E8   table cell           <NA>
#> 22        24 0112F740   table cell           <NA>
#> 23        25 4FA4104A   table cell           <NA>
#> 24        26 6411D002   table cell           <NA>
#> 25        27 4BE8A109   table cell           <NA>
#> 26        29 03C4A309   table cell           <NA>
#> 27        30 6742B786   table cell           <NA>
#> 28        31 058788C5   table cell           <NA>
#> 29        32 3B968F0F   table cell           <NA>
#> 30        33 502FD7C0   table cell           <NA>
#> 31        34 5674C826   table cell           <NA>
#> 32        35 592F4D1E   table cell           <NA>
#> 33        36 5A8BE950   table cell           <NA>
#> 34        37 448288CA   table cell           <NA>
#> 35        38 06AE66A5   table cell           <NA>
#> 36        39 22CBAAC4   table cell           <NA>
#> 37        40 7A62E4DC   table cell           <NA>
#> 38        41 3EF2CCF3   table cell           <NA>
#> 39        43 7F987E3C   table cell           <NA>
#> 40        44 182FB64A   table cell           <NA>
#> 41        45 7970FD56   table cell           <NA>
#> 42        47 5A3B1DFD   table cell           <NA>
#> 43        49 24C6DA47   table cell           <NA>
#> 44        50 7BB0A36A   table cell           <NA>
#> 45        51 6241A0F7   table cell           <NA>
#> 46        53 5D5B28B3   table cell           <NA>
#> 47        54 0D353C70   table cell           <NA>
#> 48        55 460CBC23   table cell           <NA>
#> 49        57 493EC235   table cell           <NA>
#> 50        58 07AD5415   table cell           <NA>
#> 51        59 74CD3ED9   table cell           <NA>
#> 52        60 00F0FDA0   table cell           <NA>
#> 53        61 383EA117   table cell           <NA>
#> 54        62 55B6D183   table cell           <NA>
#> 55        63 4E9DE916   table cell           <NA>
#> 56        64 3FC8AA5E   table cell           <NA>
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
#>     doc_index  para_id content_type run_index run_content_index
#> 1           1 4248496C    paragraph         1                 1
#> 2           1 4248496C    paragraph         2                 1
#> 3           2 6A358887    paragraph         1                 1
#> 4           2 6A358887    paragraph         2                 1
#> 5           2 6A358887    paragraph         3                 1
#> 6           2 6A358887    paragraph         4                 1
#> 7           2 6A358887    paragraph         5                 1
#> 8           2 6A358887    paragraph         6                 1
#> 9           2 6A358887    paragraph         7                 1
#> 10          2 6A358887    paragraph         8                 1
#> 11          2 6A358887    paragraph         9                 1
#> 12          2 6A358887    paragraph        10                 1
#> 13          2 6A358887    paragraph        11                 1
#> 14          2 6A358887    paragraph        12                 1
#> 15          2 6A358887    paragraph        13                 1
#> 16          3 22C38A25    paragraph         1                 1
#> 17          3 22C38A25    paragraph         2                 1
#> 18          4 0358E421    paragraph         1                 1
#> 19          4 0358E421    paragraph         2                 1
#> 20          5 70C2F11E    paragraph         1                 1
#> 21          5 70C2F11E    paragraph         2                 1
#> 22          5 70C2F11E    paragraph         3                 1
#> 23          5 70C2F11E    paragraph         4                 1
#> 24          5 70C2F11E    paragraph         5                 1
#> 25          5 70C2F11E    paragraph         6                 1
#> 26          5 70C2F11E    paragraph         7                 1
#> 27          6 429784E2    paragraph         1                 1
#> 28          6 429784E2    paragraph         2                 1
#> 29          7 49D1AAF7    paragraph         1                 1
#> 30          7 49D1AAF7    paragraph         2                 1
#> 31          7 49D1AAF7    paragraph         3                 1
#> 32          7 49D1AAF7    paragraph         4                 1
#> 33          8 2F68EB53    paragraph         1                 1
#> 34          8 2F68EB53    paragraph         2                 1
#> 35          9 6718A3AB    paragraph         1                 1
#> 36          9 6718A3AB    paragraph         2                 1
#> 37          9 6718A3AB    paragraph         3                 1
#> 38          9 6718A3AB    paragraph         4                 1
#> 39          9 6718A3AB    paragraph         5                 1
#> 40          9 6718A3AB    paragraph         6                 1
#> 41         10 374BD321    paragraph         1                 1
#> 42         12 0B98AFE8    paragraph         1                 1
#> 43         12 0B98AFE8    paragraph         2                 1
#> 44         12 0B98AFE8    paragraph         3                 1
#> 45         12 0B98AFE8    paragraph         4                 1
#> 46         12 0B98AFE8    paragraph         5                 1
#> 47         12 0B98AFE8    paragraph         6                 1
#> 48         12 0B98AFE8    paragraph         7                 1
#> 49         12 0B98AFE8    paragraph         8                 1
#> 50         12 0B98AFE8    paragraph         9                 1
#> 51         12 0B98AFE8    paragraph        10                 1
#> 52         13 7BA11984    paragraph         1                 1
#> 53         13 7BA11984    paragraph         2                 1
#> 54         13 7BA11984    paragraph         3                 1
#> 55         13 7BA11984    paragraph         4                 1
#> 56         14 6E5CCCAB    paragraph         1                 1
#> 57         14 6E5CCCAB    paragraph         2                 1
#> 58         14 6E5CCCAB    paragraph         3                 1
#> 59         14 6E5CCCAB    paragraph         4                 1
#> 60         14 6E5CCCAB    paragraph         5                 1
#> 61         14 6E5CCCAB    paragraph         6                 1
#> 62         14 6E5CCCAB    paragraph         7                 1
#> 63         14 6E5CCCAB    paragraph         8                 1
#> 64         14 6E5CCCAB    paragraph         9                 1
#> 65         14 6E5CCCAB    paragraph        10                 1
#> 66         14 6E5CCCAB    paragraph        11                 1
#> 67         14 6E5CCCAB    paragraph        12                 1
#> 68         14 6E5CCCAB    paragraph        13                 1
#> 69         14 6E5CCCAB    paragraph        14                 1
#> 70         14 6E5CCCAB    paragraph        15                 1
#> 71         14 6E5CCCAB    paragraph        16                 1
#> 72         14 6E5CCCAB    paragraph        17                 1
#> 73         14 6E5CCCAB    paragraph        18                 1
#> 74         14 6E5CCCAB    paragraph        19                 1
#> 75         16 105C9948   table cell         1                 1
#> 76         17 4B712A9C   table cell         1                 1
#> 77         18 4F7C7923   table cell         1                 1
#> 78         19 75402514   table cell         1                 1
#> 79         20 0AA8CF0C   table cell         1                 1
#> 80         21 07F46688   table cell         1                 1
#> 81         22 058EDA18   table cell         1                 1
#> 82         23 529A35E8   table cell         1                 1
#> 83         24 0112F740   table cell         1                 1
#> 84         25 4FA4104A   table cell         1                 1
#> 85         26 6411D002   table cell         1                 1
#> 86         27 4BE8A109   table cell         1                 1
#> 87         29 03C4A309   table cell         1                 1
#> 88         30 6742B786   table cell         1                 1
#> 89         31 058788C5   table cell         1                 1
#> 90         32 3B968F0F   table cell         1                 1
#> 91         33 502FD7C0   table cell         1                 1
#> 92         34 5674C826   table cell         1                 1
#> 93         35 592F4D1E   table cell         1                 1
#> 94         36 5A8BE950   table cell         1                 1
#> 95         37 448288CA   table cell         1                 1
#> 96         38 06AE66A5   table cell         1                 1
#> 97         39 22CBAAC4   table cell         1                 1
#> 98         40 7A62E4DC   table cell         1                 1
#> 99         41 3EF2CCF3   table cell         1                 1
#> 100        43 7F987E3C   table cell         1                 1
#> 101        44 182FB64A   table cell         1                 1
#> 102        44 182FB64A   table cell         2                 1
#> 103        44 182FB64A   table cell         3                 1
#> 104        45 7970FD56   table cell         1                 1
#> 105        47 5A3B1DFD   table cell         1                 1
#> 106        49 24C6DA47   table cell         1                 1
#> 107        50 7BB0A36A   table cell         1                 1
#> 108        51 6241A0F7   table cell         1                 1
#> 109        53 5D5B28B3   table cell         1                 1
#> 110        54 0D353C70   table cell         1                 1
#> 111        55 460CBC23   table cell         1                 1
#> 112        57 493EC235   table cell         1                 1
#> 113        58 07AD5415   table cell         1                 1
#> 114        59 74CD3ED9   table cell         1                 1
#> 115        60 00F0FDA0   table cell         1                 1
#> 116        61 383EA117   table cell         1                 1
#> 117        62 55B6D183   table cell         1                 1
#> 118        63 4E9DE916   table cell         1                 1
#> 119        64 3FC8AA5E   table cell         1                 1
#>        run_content_text image_path field_code footnote_text link
#> 1                 Title       <NA>       <NA>               <NA>
#> 2                     1       <NA>       <NA>               <NA>
#> 3          Lorem ipsum        <NA>       <NA>               <NA>
#> 4                 dolor       <NA>       <NA>               <NA>
#> 5                             <NA>       <NA>               <NA>
#> 6                   sit       <NA>       <NA>               <NA>
#> 7                             <NA>       <NA>               <NA>
#> 8                  amet       <NA>       <NA>               <NA>
#> 9                    ,        <NA>       <NA>               <NA>
#> 10          consectetur       <NA>       <NA>               <NA>
#> 11                            <NA>       <NA>               <NA>
#> 12           adipiscing       <NA>       <NA>               <NA>
#> 13                            <NA>       <NA>               <NA>
#> 14                 elit       <NA>       <NA>               <NA>
#> 15                   .        <NA>       <NA>               <NA>
#> 16                Title       <NA>       <NA>               <NA>
#> 17                    2       <NA>       <NA>               <NA>
#> 18              Quisque       <NA>       <NA>               <NA>
#> 19           tristique        <NA>       <NA>               <NA>
#> 20                    A       <NA>       <NA>               <NA>
#> 21                 ugue       <NA>       <NA>               <NA>
#> 22                            <NA>       <NA>               <NA>
#> 23                 nisi       <NA>       <NA>               <NA>
#> 24                , et        <NA>       <NA>               <NA>
#> 25            convallis       <NA>       <NA>               <NA>
#> 26                            <NA>       <NA>               <NA>
#> 27                    S       <NA>       <NA>               <NA>
#> 28   apien mollis nec.        <NA>       <NA>               <NA>
#> 29                  Sub       <NA>       <NA>               <NA>
#> 30                            <NA>       <NA>               <NA>
#> 31                title       <NA>       <NA>               <NA>
#> 32                    1       <NA>       <NA>               <NA>
#> 33              Quisque       <NA>       <NA>               <NA>
#> 34           tristique        <NA>       <NA>               <NA>
#> 35                Augue       <NA>       <NA>               <NA>
#> 36                            <NA>       <NA>               <NA>
#> 37                 nisi       <NA>       <NA>               <NA>
#> 38                , et        <NA>       <NA>               <NA>
#> 39            convallis       <NA>       <NA>               <NA>
#> 40                            <NA>       <NA>               <NA>
#> 41  Sapien mollis nec.        <NA>       <NA>               <NA>
#> 42            Phasellus       <NA>       <NA>               <NA>
#> 43      nec nunc vitae        <NA>       <NA>               <NA>
#> 44                nulla       <NA>       <NA>               <NA>
#> 45                            <NA>       <NA>               <NA>
#> 46             interdum       <NA>       <NA>               <NA>
#> 47                            <NA>       <NA>               <NA>
#> 48             volutpat       <NA>       <NA>               <NA>
#> 49                  eu        <NA>       <NA>               <NA>
#> 50                   ac       <NA>       <NA>               <NA>
#> 51              massa.        <NA>       <NA>               <NA>
#> 52                  Sub       <NA>       <NA>               <NA>
#> 53                            <NA>       <NA>               <NA>
#> 54                title       <NA>       <NA>               <NA>
#> 55                    2       <NA>       <NA>               <NA>
#> 56       Morbi rhoncus        <NA>       <NA>               <NA>
#> 57               sapien       <NA>       <NA>               <NA>
#> 58                            <NA>       <NA>               <NA>
#> 59                  sit       <NA>       <NA>               <NA>
#> 60                            <NA>       <NA>               <NA>
#> 61                 amet       <NA>       <NA>               <NA>
#> 62                            <NA>       <NA>               <NA>
#> 63                  leo       <NA>       <NA>               <NA>
#> 64                            <NA>       <NA>               <NA>
#> 65             eleifend       <NA>       <NA>               <NA>
#> 66                   ,        <NA>       <NA>               <NA>
#> 67                  vel       <NA>       <NA>               <NA>
#> 68                            <NA>       <NA>               <NA>
#> 69            fermentum       <NA>       <NA>               <NA>
#> 70                            <NA>       <NA>               <NA>
#> 71                 nisi       <NA>       <NA>               <NA>
#> 72                            <NA>       <NA>               <NA>
#> 73               mattis       <NA>       <NA>               <NA>
#> 74                   .        <NA>       <NA>               <NA>
#> 75               Petals       <NA>       <NA>               <NA>
#> 76            Internode       <NA>       <NA>               <NA>
#> 77                Sepal       <NA>       <NA>               <NA>
#> 78                Bract       <NA>       <NA>               <NA>
#> 79          5,621498349       <NA>       <NA>               <NA>
#> 80          2,462106579       <NA>       <NA>               <NA>
#> 81           18,2034091       <NA>       <NA>               <NA>
#> 82          4,994616997       <NA>       <NA>               <NA>
#> 83                   AA       <NA>       <NA>               <NA>
#> 84          2,429320759       <NA>       <NA>               <NA>
#> 85          17,65204912       <NA>       <NA>               <NA>
#> 86          4,767504884       <NA>       <NA>               <NA>
#> 87                  AAA       <NA>       <NA>               <NA>
#> 88           25,9242382       <NA>       <NA>               <NA>
#> 89          2,066051345       <NA>       <NA>               <NA>
#> 90          18,37915478       <NA>       <NA>               <NA>
#> 91          6,489375001       <NA>       <NA>               <NA>
#> 92          25,21130805       <NA>       <NA>               <NA>
#> 93          2,901582763       <NA>       <NA>               <NA>
#> 94          17,31304737       <NA>       <NA>               <NA>
#> 95          17,07215724       <NA>       <NA>               <NA>
#> 96           18,2902189       <NA>       <NA>               <NA>
#> 97            5,7858682       <NA>       <NA>               <NA>
#> 98          25,52433147       <NA>       <NA>               <NA>
#> 99          2,655642742       <NA>       <NA>               <NA>
#> 100         5,645575295       <NA>       <NA>               <NA>
#> 101              Merged       <NA>       <NA>               <NA>
#> 102                           <NA>       <NA>               <NA>
#> 103                cell       <NA>       <NA>               <NA>
#> 104         2,278691288       <NA>       <NA>               <NA>
#> 105         4,828953215       <NA>       <NA>               <NA>
#> 106         2,238467716       <NA>       <NA>               <NA>
#> 107         19,87376227       <NA>       <NA>               <NA>
#> 108         6,783500773       <NA>       <NA>               <NA>
#> 109         2,202762147       <NA>       <NA>               <NA>
#> 110         19,85326662       <NA>       <NA>               <NA>
#> 111         5,395076839       <NA>       <NA>               <NA>
#> 112         2,538375992       <NA>       <NA>               <NA>
#> 113         19,56545356       <NA>       <NA>               <NA>
#> 114         4,683617783       <NA>       <NA>               <NA>
#> 115          29,2459239       <NA>       <NA>               <NA>
#> 116         2,601945544       <NA>       <NA>               <NA>
#> 117         18,95335451       <NA>       <NA>               <NA>
#> 118                Note       <NA>       <NA>               <NA>
#> 119       New line note       <NA>       <NA>               <NA>
#>     link_to_bookmark bookmark_start character_stylename sz sz_cs
#> 1               <NA>           <NA>                <NA> NA    NA
#> 2               <NA>           <NA>                <NA> NA    NA
#> 3               <NA>          bmk_1                <NA> NA    NA
#> 4               <NA>          bmk_1                <NA> NA    NA
#> 5               <NA>          bmk_1                <NA> NA    NA
#> 6               <NA>          bmk_1                <NA> NA    NA
#> 7               <NA>          bmk_1                <NA> NA    NA
#> 8               <NA>          bmk_1                <NA> NA    NA
#> 9               <NA>          bmk_1                <NA> NA    NA
#> 10              <NA>          bmk_1                <NA> NA    NA
#> 11              <NA>          bmk_1                <NA> NA    NA
#> 12              <NA>          bmk_1                <NA> NA    NA
#> 13              <NA>          bmk_1                <NA> NA    NA
#> 14              <NA>          bmk_1                <NA> NA    NA
#> 15              <NA>          bmk_1                <NA> NA    NA
#> 16              <NA>           <NA>                <NA> NA    NA
#> 17              <NA>           <NA>                <NA> NA    NA
#> 18              <NA>           <NA>                <NA> NA    NA
#> 19              <NA>           <NA>                <NA> NA    NA
#> 20              <NA>           <NA>                <NA> NA    NA
#> 21              <NA>           <NA>                <NA> NA    NA
#> 22              <NA>           <NA>                <NA> NA    NA
#> 23              <NA>           <NA>                <NA> NA    NA
#> 24              <NA>           <NA>                <NA> NA    NA
#> 25              <NA>           <NA>                <NA> NA    NA
#> 26              <NA>           <NA>                <NA> NA    NA
#> 27              <NA>           <NA>                <NA> NA    NA
#> 28              <NA>           <NA>                <NA> NA    NA
#> 29              <NA>           <NA>                <NA> NA    NA
#> 30              <NA>           <NA>                <NA> NA    NA
#> 31              <NA>           <NA>                <NA> NA    NA
#> 32              <NA>           <NA>                <NA> NA    NA
#> 33              <NA>           <NA>                <NA> NA    NA
#> 34              <NA>           <NA>                <NA> NA    NA
#> 35              <NA>           <NA>                <NA> NA    NA
#> 36              <NA>           <NA>                <NA> NA    NA
#> 37              <NA>           <NA>                <NA> NA    NA
#> 38              <NA>           <NA>                <NA> NA    NA
#> 39              <NA>           <NA>                <NA> NA    NA
#> 40              <NA>           <NA>                <NA> NA    NA
#> 41              <NA>           <NA>                <NA> NA    NA
#> 42              <NA>           <NA>                <NA> NA    NA
#> 43              <NA>           <NA>                <NA> NA    NA
#> 44              <NA>           <NA>                <NA> NA    NA
#> 45              <NA>           <NA>                <NA> NA    NA
#> 46              <NA>           <NA>                <NA> NA    NA
#> 47              <NA>           <NA>                <NA> NA    NA
#> 48              <NA>           <NA>                <NA> NA    NA
#> 49              <NA>           <NA>                <NA> NA    NA
#> 50              <NA>           <NA>                <NA> NA    NA
#> 51              <NA>           <NA>                <NA> NA    NA
#> 52              <NA>           <NA>                <NA> NA    NA
#> 53              <NA>           <NA>                <NA> NA    NA
#> 54              <NA>           <NA>                <NA> NA    NA
#> 55              <NA>           <NA>                <NA> NA    NA
#> 56              <NA>          bmk_2                <NA> NA    NA
#> 57              <NA>          bmk_2                <NA> NA    NA
#> 58              <NA>          bmk_2                <NA> NA    NA
#> 59              <NA>          bmk_2                <NA> NA    NA
#> 60              <NA>          bmk_2                <NA> NA    NA
#> 61              <NA>          bmk_2                <NA> NA    NA
#> 62              <NA>          bmk_2                <NA> NA    NA
#> 63              <NA>          bmk_2                <NA> NA    NA
#> 64              <NA>          bmk_2                <NA> NA    NA
#> 65              <NA>          bmk_2                <NA> NA    NA
#> 66              <NA>          bmk_2                <NA> NA    NA
#> 67              <NA>          bmk_2                <NA> NA    NA
#> 68              <NA>          bmk_2                <NA> NA    NA
#> 69              <NA>          bmk_2                <NA> NA    NA
#> 70              <NA>          bmk_2                <NA> NA    NA
#> 71              <NA>          bmk_2                <NA> NA    NA
#> 72              <NA>          bmk_2                <NA> NA    NA
#> 73              <NA>          bmk_2                <NA> NA    NA
#> 74              <NA>          bmk_2                <NA> NA    NA
#> 75              <NA>           <NA>                <NA> NA    NA
#> 76              <NA>           <NA>                <NA> NA    NA
#> 77              <NA>           <NA>                <NA> NA    NA
#> 78              <NA>           <NA>                <NA> NA    NA
#> 79              <NA>           <NA>                <NA> NA    NA
#> 80              <NA>           <NA>                <NA> NA    NA
#> 81              <NA>           <NA>                <NA> NA    NA
#> 82              <NA>           <NA>                <NA> NA    NA
#> 83              <NA>           <NA>                <NA> NA    NA
#> 84              <NA>           <NA>                <NA> NA    NA
#> 85              <NA>           <NA>                <NA> NA    NA
#> 86              <NA>           <NA>                <NA> NA    NA
#> 87              <NA>           <NA>                <NA> NA    NA
#> 88              <NA>           <NA>                <NA> NA    NA
#> 89              <NA>           <NA>                <NA> NA    NA
#> 90              <NA>           <NA>                <NA> NA    NA
#> 91              <NA>           <NA>                <NA> NA    NA
#> 92              <NA>           <NA>                <NA> NA    NA
#> 93              <NA>           <NA>                <NA> NA    NA
#> 94              <NA>           <NA>                <NA> NA    NA
#> 95              <NA>           <NA>                <NA> NA    NA
#> 96              <NA>           <NA>                <NA> NA    NA
#> 97              <NA>           <NA>                <NA> NA    NA
#> 98              <NA>           <NA>                <NA> NA    NA
#> 99              <NA>           <NA>                <NA> NA    NA
#> 100             <NA>           <NA>                <NA> NA    NA
#> 101             <NA>           <NA>                <NA> NA    NA
#> 102             <NA>           <NA>                <NA> NA    NA
#> 103             <NA>           <NA>                <NA> NA    NA
#> 104             <NA>           <NA>                <NA> NA    NA
#> 105             <NA>           <NA>                <NA> NA    NA
#> 106             <NA>           <NA>                <NA> NA    NA
#> 107             <NA>           <NA>                <NA> NA    NA
#> 108             <NA>           <NA>                <NA> NA    NA
#> 109             <NA>           <NA>                <NA> NA    NA
#> 110             <NA>           <NA>                <NA> NA    NA
#> 111             <NA>           <NA>                <NA> NA    NA
#> 112             <NA>           <NA>                <NA> NA    NA
#> 113             <NA>           <NA>                <NA> NA    NA
#> 114             <NA>           <NA>                <NA> NA    NA
#> 115             <NA>           <NA>                <NA> NA    NA
#> 116             <NA>           <NA>                <NA> NA    NA
#> 117             <NA>           <NA>                <NA> NA    NA
#> 118             <NA>           <NA>                <NA> NA    NA
#> 119             <NA>           <NA>                <NA> NA    NA
#>     font_family_ascii font_family_eastasia font_family_hansi  font_family_cs
#> 1                <NA>                 <NA>              <NA>            <NA>
#> 2                <NA>                 <NA>              <NA>            <NA>
#> 3                <NA>                 <NA>              <NA>            <NA>
#> 4                <NA>                 <NA>              <NA>            <NA>
#> 5                <NA>                 <NA>              <NA>            <NA>
#> 6                <NA>                 <NA>              <NA>            <NA>
#> 7                <NA>                 <NA>              <NA>            <NA>
#> 8                <NA>                 <NA>              <NA>            <NA>
#> 9                <NA>                 <NA>              <NA>            <NA>
#> 10               <NA>                 <NA>              <NA>            <NA>
#> 11               <NA>                 <NA>              <NA>            <NA>
#> 12               <NA>                 <NA>              <NA>            <NA>
#> 13               <NA>                 <NA>              <NA>            <NA>
#> 14               <NA>                 <NA>              <NA>            <NA>
#> 15               <NA>                 <NA>              <NA>            <NA>
#> 16               <NA>                 <NA>              <NA>            <NA>
#> 17               <NA>                 <NA>              <NA>            <NA>
#> 18               <NA>                 <NA>              <NA>            <NA>
#> 19               <NA>                 <NA>              <NA>            <NA>
#> 20               <NA>                 <NA>              <NA>            <NA>
#> 21               <NA>                 <NA>              <NA>            <NA>
#> 22               <NA>                 <NA>              <NA>            <NA>
#> 23               <NA>                 <NA>              <NA>            <NA>
#> 24               <NA>                 <NA>              <NA>            <NA>
#> 25               <NA>                 <NA>              <NA>            <NA>
#> 26               <NA>                 <NA>              <NA>            <NA>
#> 27               <NA>                 <NA>              <NA>            <NA>
#> 28               <NA>                 <NA>              <NA>            <NA>
#> 29               <NA>                 <NA>              <NA>            <NA>
#> 30               <NA>                 <NA>              <NA>            <NA>
#> 31               <NA>                 <NA>              <NA>            <NA>
#> 32               <NA>                 <NA>              <NA>            <NA>
#> 33               <NA>                 <NA>              <NA>            <NA>
#> 34               <NA>                 <NA>              <NA>            <NA>
#> 35               <NA>                 <NA>              <NA>            <NA>
#> 36               <NA>                 <NA>              <NA>            <NA>
#> 37               <NA>                 <NA>              <NA>            <NA>
#> 38               <NA>                 <NA>              <NA>            <NA>
#> 39               <NA>                 <NA>              <NA>            <NA>
#> 40               <NA>                 <NA>              <NA>            <NA>
#> 41               <NA>                 <NA>              <NA>            <NA>
#> 42               <NA>                 <NA>              <NA>            <NA>
#> 43               <NA>                 <NA>              <NA>            <NA>
#> 44               <NA>                 <NA>              <NA>            <NA>
#> 45               <NA>                 <NA>              <NA>            <NA>
#> 46               <NA>                 <NA>              <NA>            <NA>
#> 47               <NA>                 <NA>              <NA>            <NA>
#> 48               <NA>                 <NA>              <NA>            <NA>
#> 49               <NA>                 <NA>              <NA>            <NA>
#> 50               <NA>                 <NA>              <NA>            <NA>
#> 51               <NA>                 <NA>              <NA>            <NA>
#> 52               <NA>                 <NA>              <NA>            <NA>
#> 53               <NA>                 <NA>              <NA>            <NA>
#> 54               <NA>                 <NA>              <NA>            <NA>
#> 55               <NA>                 <NA>              <NA>            <NA>
#> 56               <NA>                 <NA>              <NA>            <NA>
#> 57               <NA>                 <NA>              <NA>            <NA>
#> 58               <NA>                 <NA>              <NA>            <NA>
#> 59               <NA>                 <NA>              <NA>            <NA>
#> 60               <NA>                 <NA>              <NA>            <NA>
#> 61               <NA>                 <NA>              <NA>            <NA>
#> 62               <NA>                 <NA>              <NA>            <NA>
#> 63               <NA>                 <NA>              <NA>            <NA>
#> 64               <NA>                 <NA>              <NA>            <NA>
#> 65               <NA>                 <NA>              <NA>            <NA>
#> 66               <NA>                 <NA>              <NA>            <NA>
#> 67               <NA>                 <NA>              <NA>            <NA>
#> 68               <NA>                 <NA>              <NA>            <NA>
#> 69               <NA>                 <NA>              <NA>            <NA>
#> 70               <NA>                 <NA>              <NA>            <NA>
#> 71               <NA>                 <NA>              <NA>            <NA>
#> 72               <NA>                 <NA>              <NA>            <NA>
#> 73               <NA>                 <NA>              <NA>            <NA>
#> 74               <NA>                 <NA>              <NA>            <NA>
#> 75            Calibri      Times New Roman           Calibri Times New Roman
#> 76            Calibri      Times New Roman           Calibri Times New Roman
#> 77            Calibri      Times New Roman           Calibri Times New Roman
#> 78            Calibri      Times New Roman           Calibri Times New Roman
#> 79            Calibri      Times New Roman           Calibri Times New Roman
#> 80            Calibri      Times New Roman           Calibri Times New Roman
#> 81            Calibri      Times New Roman           Calibri Times New Roman
#> 82            Calibri      Times New Roman           Calibri Times New Roman
#> 83            Calibri      Times New Roman           Calibri Times New Roman
#> 84            Calibri      Times New Roman           Calibri Times New Roman
#> 85            Calibri      Times New Roman           Calibri Times New Roman
#> 86            Calibri      Times New Roman           Calibri Times New Roman
#> 87            Calibri      Times New Roman           Calibri Times New Roman
#> 88            Calibri      Times New Roman           Calibri Times New Roman
#> 89            Calibri      Times New Roman           Calibri Times New Roman
#> 90            Calibri      Times New Roman           Calibri Times New Roman
#> 91            Calibri      Times New Roman           Calibri Times New Roman
#> 92            Calibri      Times New Roman           Calibri Times New Roman
#> 93            Calibri      Times New Roman           Calibri Times New Roman
#> 94            Calibri      Times New Roman           Calibri Times New Roman
#> 95            Calibri      Times New Roman           Calibri Times New Roman
#> 96            Calibri      Times New Roman           Calibri Times New Roman
#> 97            Calibri      Times New Roman           Calibri Times New Roman
#> 98            Calibri      Times New Roman           Calibri Times New Roman
#> 99            Calibri      Times New Roman           Calibri Times New Roman
#> 100           Calibri      Times New Roman           Calibri Times New Roman
#> 101           Calibri      Times New Roman           Calibri Times New Roman
#> 102           Calibri      Times New Roman           Calibri Times New Roman
#> 103           Calibri      Times New Roman           Calibri Times New Roman
#> 104           Calibri      Times New Roman           Calibri Times New Roman
#> 105           Calibri      Times New Roman           Calibri Times New Roman
#> 106           Calibri      Times New Roman           Calibri Times New Roman
#> 107           Calibri      Times New Roman           Calibri Times New Roman
#> 108           Calibri      Times New Roman           Calibri Times New Roman
#> 109           Calibri      Times New Roman           Calibri Times New Roman
#> 110           Calibri      Times New Roman           Calibri Times New Roman
#> 111           Calibri      Times New Roman           Calibri Times New Roman
#> 112           Calibri      Times New Roman           Calibri Times New Roman
#> 113           Calibri      Times New Roman           Calibri Times New Roman
#> 114           Calibri      Times New Roman           Calibri Times New Roman
#> 115           Calibri      Times New Roman           Calibri Times New Roman
#> 116           Calibri      Times New Roman           Calibri Times New Roman
#> 117           Calibri      Times New Roman           Calibri Times New Roman
#> 118           Calibri      Times New Roman           Calibri Times New Roman
#> 119           Calibri      Times New Roman           Calibri Times New Roman
#>      bold italic underline   color shading shading_color shading_fill
#> 1   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 2   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 3   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 4   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 5   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 6   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 7   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 8   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 9   FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 10  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 11  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 12  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 13  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 14  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 15  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 16  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 17  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 18  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 19  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 20  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 21  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 22  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 23  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 24  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 25  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 26  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 27  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 28  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 29  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 30  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 31  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 32  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 33  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 34  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 35  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 36  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 37  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 38  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 39  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 40  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 41  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 42  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 43  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 44  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 45  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 46  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 47  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 48  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 49  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 50  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 51  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 52  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 53  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 54  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 55  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 56  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 57  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 58  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 59  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 60  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 61  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 62  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 63  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 64  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 65  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 66  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 67  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 68  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 69  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 70  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 71  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 72  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 73  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 74  FALSE  FALSE     FALSE    <NA>    <NA>          <NA>         <NA>
#> 75  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 76  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 77  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 78  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 79  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 80  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 81  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 82  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 83  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 84  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 85  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 86  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 87  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 88  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 89  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 90  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 91  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 92  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 93  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 94  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 95  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 96  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 97  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 98  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 99  FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 100 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 101 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 102 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 103 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 104 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 105 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 106 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 107 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 108 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 109 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 110 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 111 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 112 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 113 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 114 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 115 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 116 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 117 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 118 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#> 119 FALSE  FALSE     FALSE #000000    <NA>          <NA>         <NA>
#>     paragraph_stylename keep_with_next  align level num_id table_index row_id
#> 1             heading 1          FALSE   <NA>    NA     NA          NA     NA
#> 2             heading 1          FALSE   <NA>    NA     NA          NA     NA
#> 3                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 4                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 5                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 6                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 7                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 8                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 9                  <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 10                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 11                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 12                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 13                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 14                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 15                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 16            heading 1          FALSE   <NA>    NA     NA          NA     NA
#> 17            heading 1          FALSE   <NA>    NA     NA          NA     NA
#> 18       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 19       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 20       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 21       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 22       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 23       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 24       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 25       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 26       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 27       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 28       List Paragraph          FALSE   <NA>     1      2          NA     NA
#> 29            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 30            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 31            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 32            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 33       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 34       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 35       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 36       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 37       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 38       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 39       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 40       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 41       List Paragraph          FALSE   <NA>     1      1          NA     NA
#> 42                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 43                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 44                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 45                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 46                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 47                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 48                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 49                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 50                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 51                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 52            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 53            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 54            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 55            heading 2          FALSE   <NA>    NA     NA          NA     NA
#> 56                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 57                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 58                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 59                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 60                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 61                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 62                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 63                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 64                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 65                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 66                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 67                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 68                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 69                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 70                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 71                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 72                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 73                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 74                 <NA>          FALSE   <NA>    NA     NA          NA     NA
#> 75                 <NA>          FALSE   <NA>    NA     NA           1      1
#> 76                 <NA>          FALSE   <NA>    NA     NA           1      1
#> 77                 <NA>          FALSE   <NA>    NA     NA           1      1
#> 78                 <NA>          FALSE   <NA>    NA     NA           1      1
#> 79                 <NA>          FALSE  right    NA     NA           1      2
#> 80                 <NA>          FALSE  right    NA     NA           1      2
#> 81                 <NA>          FALSE  right    NA     NA           1      2
#> 82                 <NA>          FALSE  right    NA     NA           1      3
#> 83                 <NA>          FALSE  right    NA     NA           1      3
#> 84                 <NA>          FALSE  right    NA     NA           1      3
#> 85                 <NA>          FALSE  right    NA     NA           1      3
#> 86                 <NA>          FALSE  right    NA     NA           1      4
#> 87                 <NA>          FALSE center    NA     NA           1      4
#> 88                 <NA>          FALSE  right    NA     NA           1      5
#> 89                 <NA>          FALSE  right    NA     NA           1      5
#> 90                 <NA>          FALSE  right    NA     NA           1      5
#> 91                 <NA>          FALSE  right    NA     NA           1      6
#> 92                 <NA>          FALSE  right    NA     NA           1      6
#> 93                 <NA>          FALSE  right    NA     NA           1      6
#> 94                 <NA>          FALSE  right    NA     NA           1      6
#> 95                 <NA>          FALSE  right    NA     NA           1      6
#> 96                 <NA>          FALSE  right    NA     NA           1      6
#> 97                 <NA>          FALSE  right    NA     NA           1      7
#> 98                 <NA>          FALSE  right    NA     NA           1      7
#> 99                 <NA>          FALSE  right    NA     NA           1      7
#> 100                <NA>          FALSE  right    NA     NA           1      8
#> 101                <NA>          FALSE  right    NA     NA           1      8
#> 102                <NA>          FALSE  right    NA     NA           1      8
#> 103                <NA>          FALSE  right    NA     NA           1      8
#> 104                <NA>          FALSE  right    NA     NA           1      8
#> 105                <NA>          FALSE  right    NA     NA           1      9
#> 106                <NA>          FALSE  right    NA     NA           1      9
#> 107                <NA>          FALSE  right    NA     NA           1      9
#> 108                <NA>          FALSE  right    NA     NA           1     10
#> 109                <NA>          FALSE  right    NA     NA           1     10
#> 110                <NA>          FALSE  right    NA     NA           1     10
#> 111                <NA>          FALSE  right    NA     NA           1     11
#> 112                <NA>          FALSE  right    NA     NA           1     11
#> 113                <NA>          FALSE  right    NA     NA           1     11
#> 114                <NA>          FALSE  right    NA     NA           1     12
#> 115                <NA>          FALSE  right    NA     NA           1     12
#> 116                <NA>          FALSE  right    NA     NA           1     12
#> 117                <NA>          FALSE  right    NA     NA           1     12
#> 118                <NA>          FALSE  right    NA     NA           1     13
#> 119                <NA>          FALSE  right    NA     NA           1     13
#>     cell_id col_span row_span is_header table_stylename
#> 1        NA     <NA>       NA        NA            <NA>
#> 2        NA     <NA>       NA        NA            <NA>
#> 3        NA     <NA>       NA        NA            <NA>
#> 4        NA     <NA>       NA        NA            <NA>
#> 5        NA     <NA>       NA        NA            <NA>
#> 6        NA     <NA>       NA        NA            <NA>
#> 7        NA     <NA>       NA        NA            <NA>
#> 8        NA     <NA>       NA        NA            <NA>
#> 9        NA     <NA>       NA        NA            <NA>
#> 10       NA     <NA>       NA        NA            <NA>
#> 11       NA     <NA>       NA        NA            <NA>
#> 12       NA     <NA>       NA        NA            <NA>
#> 13       NA     <NA>       NA        NA            <NA>
#> 14       NA     <NA>       NA        NA            <NA>
#> 15       NA     <NA>       NA        NA            <NA>
#> 16       NA     <NA>       NA        NA            <NA>
#> 17       NA     <NA>       NA        NA            <NA>
#> 18       NA     <NA>       NA        NA            <NA>
#> 19       NA     <NA>       NA        NA            <NA>
#> 20       NA     <NA>       NA        NA            <NA>
#> 21       NA     <NA>       NA        NA            <NA>
#> 22       NA     <NA>       NA        NA            <NA>
#> 23       NA     <NA>       NA        NA            <NA>
#> 24       NA     <NA>       NA        NA            <NA>
#> 25       NA     <NA>       NA        NA            <NA>
#> 26       NA     <NA>       NA        NA            <NA>
#> 27       NA     <NA>       NA        NA            <NA>
#> 28       NA     <NA>       NA        NA            <NA>
#> 29       NA     <NA>       NA        NA            <NA>
#> 30       NA     <NA>       NA        NA            <NA>
#> 31       NA     <NA>       NA        NA            <NA>
#> 32       NA     <NA>       NA        NA            <NA>
#> 33       NA     <NA>       NA        NA            <NA>
#> 34       NA     <NA>       NA        NA            <NA>
#> 35       NA     <NA>       NA        NA            <NA>
#> 36       NA     <NA>       NA        NA            <NA>
#> 37       NA     <NA>       NA        NA            <NA>
#> 38       NA     <NA>       NA        NA            <NA>
#> 39       NA     <NA>       NA        NA            <NA>
#> 40       NA     <NA>       NA        NA            <NA>
#> 41       NA     <NA>       NA        NA            <NA>
#> 42       NA     <NA>       NA        NA            <NA>
#> 43       NA     <NA>       NA        NA            <NA>
#> 44       NA     <NA>       NA        NA            <NA>
#> 45       NA     <NA>       NA        NA            <NA>
#> 46       NA     <NA>       NA        NA            <NA>
#> 47       NA     <NA>       NA        NA            <NA>
#> 48       NA     <NA>       NA        NA            <NA>
#> 49       NA     <NA>       NA        NA            <NA>
#> 50       NA     <NA>       NA        NA            <NA>
#> 51       NA     <NA>       NA        NA            <NA>
#> 52       NA     <NA>       NA        NA            <NA>
#> 53       NA     <NA>       NA        NA            <NA>
#> 54       NA     <NA>       NA        NA            <NA>
#> 55       NA     <NA>       NA        NA            <NA>
#> 56       NA     <NA>       NA        NA            <NA>
#> 57       NA     <NA>       NA        NA            <NA>
#> 58       NA     <NA>       NA        NA            <NA>
#> 59       NA     <NA>       NA        NA            <NA>
#> 60       NA     <NA>       NA        NA            <NA>
#> 61       NA     <NA>       NA        NA            <NA>
#> 62       NA     <NA>       NA        NA            <NA>
#> 63       NA     <NA>       NA        NA            <NA>
#> 64       NA     <NA>       NA        NA            <NA>
#> 65       NA     <NA>       NA        NA            <NA>
#> 66       NA     <NA>       NA        NA            <NA>
#> 67       NA     <NA>       NA        NA            <NA>
#> 68       NA     <NA>       NA        NA            <NA>
#> 69       NA     <NA>       NA        NA            <NA>
#> 70       NA     <NA>       NA        NA            <NA>
#> 71       NA     <NA>       NA        NA            <NA>
#> 72       NA     <NA>       NA        NA            <NA>
#> 73       NA     <NA>       NA        NA            <NA>
#> 74       NA     <NA>       NA        NA            <NA>
#> 75        1        1        1      TRUE   Light Shading
#> 76        2        1        1      TRUE   Light Shading
#> 77        3        1        1      TRUE   Light Shading
#> 78        4        1        1      TRUE   Light Shading
#> 79        1        2        1     FALSE   Light Shading
#> 80        3        2        1     FALSE   Light Shading
#> 81        4        2        1     FALSE   Light Shading
#> 82        1        1        1     FALSE   Light Shading
#> 83        2        1        2     FALSE   Light Shading
#> 84        3        1        1     FALSE   Light Shading
#> 85        4        1        1     FALSE   Light Shading
#> 86        1        1        1     FALSE   Light Shading
#> 87        3        2        1     FALSE   Light Shading
#> 88        1        2        1     FALSE   Light Shading
#> 89        3        1        1     FALSE   Light Shading
#> 90        4        1        1     FALSE   Light Shading
#> 91        1        1        1     FALSE   Light Shading
#> 92        2        1        1     FALSE   Light Shading
#> 93        3        1        1     FALSE   Light Shading
#> 94        4        1        1     FALSE   Light Shading
#> 95        4        1        1     FALSE   Light Shading
#> 96        4        1        3     FALSE   Light Shading
#> 97        1        1        1     FALSE   Light Shading
#> 98        2        1        1     FALSE   Light Shading
#> 99        3        1        1     FALSE   Light Shading
#> 100       1        1        1     FALSE   Light Shading
#> 101       2        1        4     FALSE   Light Shading
#> 102       2        1        4     FALSE   Light Shading
#> 103       2        1        4     FALSE   Light Shading
#> 104       3        1        1     FALSE   Light Shading
#> 105       1        1        1     FALSE   Light Shading
#> 106       3        1        1     FALSE   Light Shading
#> 107       4        1        1     FALSE   Light Shading
#> 108       1        1        1     FALSE   Light Shading
#> 109       3        1        1     FALSE   Light Shading
#> 110       4        1        1     FALSE   Light Shading
#> 111       1        1        1     FALSE   Light Shading
#> 112       3        1        1     FALSE   Light Shading
#> 113       4        1        1     FALSE   Light Shading
#> 114       1        1        1     FALSE   Light Shading
#> 115       2        1        1     FALSE   Light Shading
#> 116       3        1        1     FALSE   Light Shading
#> 117       4        1        1     FALSE   Light Shading
#> 118       1        4        1     FALSE   Light Shading
#> 119       4        4        1     FALSE   Light Shading
```
