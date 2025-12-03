# External image

Wraps an image in an object that can then be embedded in a PowerPoint
slide or within a Word paragraph.

The image is added as a shape in PowerPoint (it is not possible to mix
text and images in a PowerPoint form). With a Word document, the image
will be added inside a paragraph.

## Usage

``` r
external_img(
  src,
  width = 0.5,
  height = 0.2,
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

- unit:

  unit for width and height, one of "in", "cm", "mm".

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

[ph_with](https://davidgohel.github.io/officer/dev/reference/ph_with.md),
[body_add](https://davidgohel.github.io/officer/dev/reference/body_add.md),
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md)

Other run functions for reporting:
[`floating_external_img()`](https://davidgohel.github.io/officer/dev/reference/floating_external_img.md),
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
# wrap r logo with external_img ----
srcfile <- file.path(R.home("doc"), "html", "logo.jpg")
extimg <- external_img(
  src = srcfile, height = 1.06 / 2,
  width = 1.39 / 2
)

# pptx example ----
doc <- read_pptx()
doc <- add_slide(doc, "Title and Content")
doc <- ph_with(
  x = doc, value = extimg,
  location = ph_location_type(type = "body"),
  use_loc_size = FALSE
)
print(doc, target = tempfile(fileext = ".pptx"))

fp_t <- fp_text(font.size = 20, color = "red")
an_fpar <- fpar(extimg, ftext(" is cool!", fp_t))

# docx example ----
x <- read_docx()
x <- body_add(x, an_fpar)
print(x, target = tempfile(fileext = ".docx"))
```
