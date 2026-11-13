# Equation

Create an equation written in LaTeX, to be used in a paragraph made with
[`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md).

The equation is converted by the package 'equatags' into an 'Office'
equation ('Office Open XML Math'): it stays editable in 'Word' and
'PowerPoint'. Use it with
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md)
in a 'Word' document, with
[`ph_with()`](https://davidgohel.github.io/officer/dev/reference/ph_with.md)
in a 'PowerPoint' presentation and with
[`sheet_add_drawing()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.md)
in an 'Excel' workbook, where equations are in a text box ('Excel' does
not accept equations in cells).

In HTML
([`to_html()`](https://davidgohel.github.io/officer/dev/reference/to_html.md)),
equations are drawn by 'KaTeX' (its stylesheet is linked). In RTF
documents
([`rtf_add()`](https://davidgohel.github.io/officer/dev/reference/rtf_add.md)),
they are images drawn by 'equatags' (R \>= 4.3), placed on the baseline
of the text; with an older version of R, the LaTeX source is written as
text, with a warning.

## Usage

``` r
run_equation(x, prop = NULL)
```

## Arguments

- x:

  the equation, a LaTeX string. The functions supported by 'KaTeX' can
  be used, see <https://katex.org/docs/supported.html>.

- prop:

  formatting text properties returned by
  [`fp_text()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md)
  or
  [`fp_text_lite()`](https://davidgohel.github.io/officer/dev/reference/fp_text.md).
  Only the font size and the color are used, the font is the math font
  of the document ('Cambria Math' by default). With `NULL`, the default
  size and color of the document or of the shape.

## Inline and display style

As in 'Word', the style of the equation depends on its paragraph:

- an equation alone in its paragraph uses the display style (larger
  fractions, limits above and below sums) and is aligned with the
  alignment of the paragraph (argument `text.align` of
  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md));

- an equation in a paragraph that contains something else, even a space,
  uses the inline style.

## usage

You can use this function in conjunction with
[fpar](https://davidgohel.github.io/officer/dev/reference/fpar.md) to
create paragraphs consisting of differently formatted text parts. You
can also use this function as an *r chunk* in an R Markdown document
made with package officedown.

## See also

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
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
library(officer)

quadratic <- "x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}"
euler <- "e^{i\\pi} + 1 = 0"

doc <- read_docx()
# an equation alone in its paragraph: display style, centered
doc <- body_add_fpar(
  doc,
  fpar(run_equation(quadratic), fp_p = fp_par(text.align = "center"))
)
# an equation in a sentence: inline style, 14 points, in red
doc <- body_add_fpar(
  doc,
  fpar(
    "Euler's identity: ",
    run_equation(euler, prop = fp_text_lite(font.size = 14, color = "red")),
    "."
  )
)
print(doc, target = tempfile(fileext = ".docx"))

pres <- read_pptx()
pres <- add_slide(pres, layout = "Title and Content")
pres <- ph_with(
  pres,
  block_list(
    fpar("The quadratic formula:"),
    fpar(run_equation(quadratic), fp_p = fp_par(text.align = "center"))
  ),
  location = ph_location_type("body")
)
print(pres, target = tempfile(fileext = ".pptx"))

wb <- read_xlsx()
wb <- sheet_add_drawing(
  wb,
  value = fpar(run_equation(quadratic)),
  sheet = "Feuil1",
  anchor = "C2"
)
print(wb, target = tempfile(fileext = ".xlsx"))
```
