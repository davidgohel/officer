# Formatted paragraph

Create a paragraph representation by concatenating formatted text or
images. The result can be inserted in a Word document or a PowerPoint
presentation and can also be inserted in a
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)
call.

All its arguments will be concatenated to create a paragraph where
chunks of text and images are associated with formatting properties.

`fpar()` supports
[`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md),
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
`run_*()` functions (i.e.
[`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md))
when output is Word, and simple strings.

Default text and paragraph formatting properties can also be modified
with function [`update()`](https://rdrr.io/r/stats/update.html).

## Usage

``` r
fpar(
  ...,
  fp_p = fp_par(word_style = NA_character_),
  fp_t = fp_text_lite(),
  values = NULL
)

# S3 method for class 'fpar'
update(object, fp_p = NULL, fp_t = NULL, ...)
```

## Arguments

- ...:

  cot objects
  ([`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md),
  [`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md))

- fp_p:

  paragraph formatting properties, see
  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)

- fp_t:

  default text formatting properties. This is used as text formatting
  properties when simple text is provided as argument, see
  [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md).

- values:

  a list of cot objects. If provided, argument `...` will be ignored.

- object:

  fpar object

## See also

[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md),
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)

Other block functions for reporting:
[`block_caption()`](https://davidgohel.github.io/officer/dev/reference/block_caption.md),
[`block_gg()`](https://davidgohel.github.io/officer/dev/reference/block_gg.md),
[`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md),
[`block_pour_docx()`](https://davidgohel.github.io/officer/dev/reference/block_pour_docx.md),
[`block_section()`](https://davidgohel.github.io/officer/dev/reference/block_section.md),
[`block_table()`](https://davidgohel.github.io/officer/dev/reference/block_table.md),
[`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md),
[`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
[`unordered_list()`](https://davidgohel.github.io/officer/dev/reference/unordered_list.md)

## Examples

``` r
fpar(ftext("hello", shortcuts$fp_bold()))
#> $chunks
#> $chunks[[1]]
#> text: hello
#> format:
#>   font.size italic bold underlined strike color     shading fontname
#> 1        10  FALSE TRUE      FALSE  FALSE black transparent    Arial
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

# mix text and image -----
img.file <- file.path( R.home("doc"), "html", "logo.jpg" )

bold_face <- shortcuts$fp_bold(font.size = 12)
bold_redface <- update(bold_face, color = "red")
fpar_1 <- fpar(
  "Hello World, ",
  ftext("how ", prop = bold_redface ),
  external_img(src = img.file, height = 1.06/2, width = 1.39/2),
  ftext(" you?", prop = bold_face ) )
fpar_1
#> $chunks
#> $chunks[[1]]
#> [1] "Hello World, "
#> 
#> $chunks[[2]]
#> text: how 
#> format:
#>   font.size italic bold underlined strike color     shading fontname
#> 1        12  FALSE TRUE      FALSE  FALSE   red transparent    Arial
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1       Arial             Arial          Arial       baseline
#> 
#> $chunks[[3]]
#> [1] "/opt/R/4.5.2/lib/R/doc/html/logo.jpg"
#> attr(,"class")
#> [1] "external_img" "cot"          "run"         
#> attr(,"dims")
#> attr(,"dims")$width
#> [1] 0.695
#> 
#> attr(,"dims")$height
#> [1] 0.53
#> 
#> attr(,"alt")
#> [1] ""
#> 
#> $chunks[[4]]
#> text:  you?
#> format:
#>   font.size italic bold underlined strike color     shading fontname
#> 1        12  FALSE TRUE      FALSE  FALSE black transparent    Arial
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

img_in_par <- fpar(
  external_img(src = img.file, height = 1.06/2, width = 1.39/2),
  fp_p = fp_par(text.align = "center") )
```
