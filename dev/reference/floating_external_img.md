# Floating external image

Wraps an image in an object that can be embedded as a floating image in
a 'Word' document. Unlike
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
which creates inline images, this function creates floating images that
can be positioned anywhere on the page and allow text wrapping around
them.

## Usage

``` r
floating_external_img(
  src,
  width = 0.5,
  height = 0.2,
  pos_x = 0,
  pos_y = 0,
  pos_h_from = "margin",
  pos_v_from = "margin",
  wrap_type = "square",
  wrap_side = "bothSides",
  wrap_dist_top = 0,
  wrap_dist_bottom = 0,
  wrap_dist_left = 0.125,
  wrap_dist_right = 0.125,
  unit = "in",
  guess_size = FALSE,
  alt = ""
)
```

## Arguments

- src:

  image file path

- width, height:

  size of the image file. It can be ignored if parameter
  `guess_size=TRUE`, see parameter `guess_size`.

- pos_x, pos_y:

  horizontal and vertical position of the image relative to the anchor
  point

- pos_h_from:

  horizontal positioning reference point, one of "margin", "page",
  "column", "character"

- pos_v_from:

  vertical positioning reference point, one of "margin", "page",
  "paragraph", "line"

- wrap_type:

  text wrapping type, one of "square", "topAndBottom", "through",
  "tight", "none"

- wrap_side:

  which side text wraps around, one of "bothSides", "left", "right",
  "largest"

- wrap_dist_top, wrap_dist_bottom, wrap_dist_left, wrap_dist_right:

  distance between image and text (in inches)

- unit:

  unit for width, height, pos_x and pos_y, one of "in", "cm", "mm".

- guess_size:

  If package 'magick' is installed, this option can be used (set it to
  `TRUE`). The images will be read and width and height will be guessed.

- alt:

  alternative text for images

## usage

You can use this function in conjunction with
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md) to
create paragraphs consisting of differently formatted text parts. You
can also use this function as an *r chunk* in an R Markdown document
made with package officedown.

## See also

[external_img](https://davidgohel.github.io/officer/dev/reference/external_img.md),
[body_add](https://davidgohel.github.io/officer/dev/reference/body_add.md),
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md),
[rtf_doc](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md),
[rtf_add](https://davidgohel.github.io/officer/dev/reference/rtf_add.md)

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
[`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/dev/reference/hyperlink_ftext.md),
[`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/dev/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/dev/reference/run_comment.md),
[`run_footnote()`](https://davidgohel.github.io/officer/dev/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/dev/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md),
[`run_wordtext()`](https://davidgohel.github.io/officer/dev/reference/run_wordtext.md)

## Examples

``` r
library(officer)
srcfile <- file.path(R.home("doc"), "html", "logo.jpg")
floatimg <- floating_external_img(
  src = srcfile, height = 1.06 / 2, width = 1.39 / 2,
  pos_x = 0, pos_y = 0,
  pos_h_from = "margin", pos_v_from = "margin"
)

text <- paste0(
  " is a floating image in a ",
  paste0(rep("very ", 30), collapse = ""),
  " long text!"
)

# docx example ----
x <- read_docx()
fp_t <- fp_text(font.size = 20, color = "red")
an_fpar <- fpar(floatimg, ftext(text, fp_t))
x <- body_add_fpar(x, an_fpar)
print(x, target = tempfile(fileext = ".docx"))

# rtf example ----
rtf_doc <- rtf_doc()
rtf_doc <- rtf_add(rtf_doc, an_fpar)
print(rtf_doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpMgcMop/file179a1e22716a.rtf"
```
