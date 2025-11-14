# List of blocks

A list of blocks can be used to gather several blocks (paragraphs,
tables, ...) into a single object. The result can be added into a Word
document or a PowerPoint presentation.

## Usage

``` r
block_list(...)
```

## Arguments

- ...:

  a list of blocks. When output is only for Word, objects of class
  [`external_img()`](https://davidgohel.github.io/officer/reference/external_img.md)
  can also be used in fpar construction to mix text and images in a
  single paragraph. Supported objects are:
  [`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md),
  [`block_pour_docx()`](https://davidgohel.github.io/officer/reference/block_pour_docx.md),
  [`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
  [`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md),
  [`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
  [`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
  [`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md).

## See also

[`ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.md),
[`body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md)

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/reference/block_gg.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/reference/block_toc.md),
[`fpar()`](https://davidgohel.github.io/officer/reference/fpar.md),
[`plot_instr()`](https://davidgohel.github.io/officer/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/reference/unordered_list.md)

## Examples

``` r
# block list ------

img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
fpt_blue_bold <- fp_text(color = "#006699", bold = TRUE)
fpt_red_italic <- fp_text(color = "#C32900", italic = TRUE)


## This can be only be used in a MS word output as pptx does
## not support paragraphs made of text and images.
## (actually it can be used but image will not appear in the
## pptx output)
value <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(ftext("hello", fpt_blue_bold), " ",
       ftext("world", fpt_red_italic)),
  fpar(
    ftext("hello world", fpt_red_italic),
          external_img(
            src = img.file, height = 1.06, width = 1.39)))
value
#> [[1]]
#> $chunks
#> $chunks[[1]]
#> text: hello world
#> format:
#>   font.size italic bold underlined strike   color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE #006699 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> [[2]]
#> $chunks
#> $chunks[[1]]
#> text: hello
#> format:
#>   font.size italic bold underlined strike   color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE #006699 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> $chunks[[2]]
#> [1] " "
#> 
#> $chunks[[3]]
#> text: world
#> format:
#>   font.size italic  bold underlined strike   color     shading fontname
#> 1        10   TRUE FALSE      FALSE  FALSE #C32900 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> [[3]]
#> $chunks
#> $chunks[[1]]
#> text: hello world
#> format:
#>   font.size italic  bold underlined strike   color     shading fontname
#> 1        10   TRUE FALSE      FALSE  FALSE #C32900 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> $chunks[[2]]
#> [1] "/opt/R/4.5.2/lib/R/doc/html/logo.jpg"
#> attr(,"class")
#> [1] "external_img" "cot"          "run"         
#> attr(,"dims")
#> attr(,"dims")$width
#> [1] 1.39
#> 
#> attr(,"dims")$height
#> [1] 1.06
#> 
#> attr(,"alt")
#> [1] ""
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> attr(,"class")
#> [1] "block_list" "block"     

doc <- read_docx()
doc <- body_add(doc, value)
print(doc, target = tempfile(fileext = ".docx"))


value <- block_list(
  fpar(ftext("hello world", fpt_blue_bold)),
  fpar(ftext("hello", fpt_blue_bold), " ",
       ftext("world", fpt_red_italic)),
  fpar(
    ftext("blah blah blah", fpt_red_italic)))
value
#> [[1]]
#> $chunks
#> $chunks[[1]]
#> text: hello world
#> format:
#>   font.size italic bold underlined strike   color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE #006699 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> [[2]]
#> $chunks
#> $chunks[[1]]
#> text: hello
#> format:
#>   font.size italic bold underlined strike   color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE #006699 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> $chunks[[2]]
#> [1] " "
#> 
#> $chunks[[3]]
#> text: world
#> format:
#>   font.size italic  bold underlined strike   color     shading fontname
#> 1        10   TRUE FALSE      FALSE  FALSE #C32900 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> [[3]]
#> $chunks
#> $chunks[[1]]
#> text: blah blah blah
#> format:
#>   font.size italic  bold underlined strike   color     shading fontname
#> 1        10   TRUE FALSE      FALSE  FALSE #C32900 transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> 
#> $fp_p
#>                     values
#> text.align            left
#> padding.top              0
#> padding.bottom           0
#> padding.left             0
#> padding.right            0
#> shading.color  transparent
#> borders:
#>        width color style
#> top        0 black solid
#> bottom     0 black solid
#> left       0 black solid
#> right      0 black solid
#> 
#> $fp_t
#>   font.size italic bold underlined strike color shading fontname fontname_cs
#> 1        NA     NA   NA         NA     NA    NA      NA       NA          NA
#>   fontname_eastasia fontname.hansi vertical_align
#> 1                NA             NA       baseline
#> 
#> attr(,"class")
#> [1] "fpar"  "block"
#> 
#> attr(,"class")
#> [1] "block_list" "block"     

doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(doc, value, location = ph_location_type(type = "body"))
print(doc, target = tempfile(fileext = ".pptx"))
```
