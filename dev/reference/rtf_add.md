# Add content into an RTF document

This function add 'officer' objects into an RTF document. Values are
added as new paragraphs. See section 'Methods (by class)' that list
supported objects.

## Usage

``` r
rtf_add(x, value, ...)

# S3 method for class 'block_section'
rtf_add(x, value, ...)

# S3 method for class 'character'
rtf_add(x, value, style = NULL, ...)

# S3 method for class 'factor'
rtf_add(x, value, style = NULL, ...)

# S3 method for class 'double'
rtf_add(x, value, formatter = formatC, style = NULL, ...)

# S3 method for class 'fpar'
rtf_add(x, value, style = NULL, ...)

# S3 method for class 'block_list'
rtf_add(x, value, style = NULL, ...)

# S3 method for class 'block_toc'
rtf_add(x, value, ...)

# S3 method for class 'gg'
rtf_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)

# S3 method for class 'plot_instr'
rtf_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)
```

## Arguments

- x:

  rtf object, created by
  [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md).

- value:

  object to add in the document. Supported objects are vectors,
  graphics, block of formatted paragraphs. Use package 'flextable' to
  add tables.

- ...:

  further arguments passed to or from other methods. When adding a
  `ggplot` object or
  [plot_instr](https://davidgohel.github.io/officer/dev/reference/plot_instr.md),
  these arguments will be used by png function. See section 'Methods' to
  see what arguments can be used.

- style:

  style identifier (`style_id`) of a paragraph style registered on the
  document. Defaults to `NULL` (use the document's normal style). See
  [`rtf_set_paragraph_style()`](https://davidgohel.github.io/officer/dev/reference/rtf_set_paragraph_style.md)
  and
  [`rtf_styles_info()`](https://davidgohel.github.io/officer/dev/reference/rtf_styles_info.md).

- formatter:

  function used to format the numerical values

- width:

  height in inches

- height:

  height in inches

- res:

  resolution of the png image in ppi

- scale:

  Multiplicative scaling factor, same as in ggsave

- ppr:

  [`fp_par()`](https://davidgohel.github.io/officer/dev/reference/fp_par.md)
  to apply to paragraph.

## Methods (by class)

- `rtf_add(block_section)`: add a new section definition

- `rtf_add(character)`: add characters as new paragraphs

- `rtf_add(factor)`: add a factor vector as new paragraphs

- `rtf_add(double)`: add a double vector as new paragraphs

- `rtf_add(fpar)`: add an
  [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)

- `rtf_add(block_list)`: add an
  [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)

- `rtf_add(block_toc)`: add a
  [`block_toc()`](https://davidgohel.github.io/officer/dev/reference/block_toc.md)
  (table of contents). Word populates the TOC at open time from
  paragraphs styled with the built-in heading styles (which carry an
  outline level). LibreOffice will not render the TOC automatically.

- `rtf_add(gg)`: add a ggplot2

- `rtf_add(plot_instr)`: add a
  [`plot_instr()`](https://davidgohel.github.io/officer/dev/reference/plot_instr.md)
  object

## Section model in RTF

A `block_section` added with `rtf_add()` applies to the content that
**follows** the call: it opens a new section whose layout (orientation,
columns, margins, headers / footers) is inherited by every paragraph,
table or graphic added afterwards, until the next `block_section` (or
the end of the document).

Typical pattern: declare the section, then add the content.

    doc <- rtf_doc() |>
      rtf_add(block_section(prop_section(
        page_size = page_size(orient = "landscape")
      ))) |>
      rtf_add(fpar("This paragraph is in landscape orientation."))

The first section of the document is configured via the `def_sec`
argument of
[`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md)
(a
[prop_section](https://davidgohel.github.io/officer/dev/reference/prop_section.md)
object, not a `block_section`). It applies to every element added before
the first `rtf_add(block_section(...))` call.

The Word output uses the opposite model:
[`body_end_block_section()`](https://davidgohel.github.io/officer/dev/reference/body_end_block_section.md)
applies to the content that *precedes* the call. See
[`body_end_block_section()`](https://davidgohel.github.io/officer/dev/reference/body_end_block_section.md).

## Examples

``` r
## Simple RTF example ----

library(officer)

def_text <- fp_text_lite(color = "#006699", bold = TRUE)
center_par <- fp_par(text.align = "center", padding = 3)

doc <- rtf_doc(
  normal_par = fp_par(line_spacing = 1.4, padding = 3)
)

doc <- rtf_add(
  x = doc,
  value = fpar(
    ftext("how are you?", prop = def_text),
    fp_p = fp_par(text.align = "center")
  )
)

a_paragraph <- fpar(
  ftext("Here is a date: ", prop = def_text),
  run_word_field(field = "Date \\@ \"MMMM d yyyy\""),
  fp_p = center_par
)
doc <- rtf_add(
  x = doc,
  value = block_list(
    a_paragraph,
    a_paragraph,
    a_paragraph
  )
)

if (require("ggplot2")) {
  gg <- gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  doc <- rtf_add(doc, gg, width = 3, height = 4, ppr = center_par)
}
anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
})
doc <- rtf_add(doc, anyplot, width = 5, height = 4, ppr = center_par)

print(doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpZCXWQh/file18db1ee9687e.rtf"


## RTF example with sections ----

library(officer)

quick_section_header <- function(label) {
  block_list(
    fpar(
      ftext(label, fp_text_lite(bold = TRUE, color = "#006699"))
    )
  )
}

quick_section_footer <- function(label) {
  block_list(
    fpar(
      "Page ",
      run_word_field(field = "PAGE  \\* MERGEFORMAT")
    )
  )
}

quick_hello_world <- function(doc) {
  rtf_add(
    doc,
    fpar(
      ftext(
        "Hello World"
      ),
      fp_p = fp_par(text.align = "left")
    )
  )
}

three_cols_section <- block_section(
  prop_section(
    type = "continuous",
    section_columns = section_columns(widths = c(1.7, 1.7, 1.7), space = 0.25),
    header_default = quick_section_header("Three columns section"),
    footer_default = quick_section_footer("Three columns section")
  )
)

doc <- rtf_doc(
  def_sec = prop_section(
    header_default = quick_section_header("Default section"),
    footer_default = quick_section_footer("Default section")
  ),
  normal_par = fp_par(padding = 3)
)
doc <- rtf_add(doc, block_toc(level = 3))

doc <- rtf_add(doc, "Sections demo", style = "heading 1")
doc <- rtf_add(doc, "Default section", style = "heading 2")
doc <- quick_hello_world(doc)

if (require("ggplot2")) {
  gg_iris <- ggplot(iris, aes(Sepal.Length, Petal.Length, colour = Species)) +
    geom_point() +
    theme_minimal()
  doc <- rtf_add(doc, gg_iris, width = 4, height = 3)
}

doc <- rtf_add(doc, "Three columns", style = "heading 2")

doc <- rtf_add(doc, three_cols_section)

titles_list <- block_list(
  fpar(ftext("Left Title"), fp_p = fp_par(text.align = "left")),
  fpar(
    run_columnbreak(),
    ftext("Centered Title"),
    fp_p = fp_par(text.align = "center")
  ),
  fpar(
    run_columnbreak(),
    ftext("Right Title"),
    fp_p = fp_par(text.align = "right")
  )
)
doc <- rtf_add(doc, titles_list)
doc <- rtf_add(doc, block_section(prop_section()))
doc <- rtf_add(doc, fpar(run_linebreak()))
doc <- quick_hello_world(doc)

landscape_section <- block_section(prop_section(
  type = "nextPage",
  page_size = page_size(orient = "landscape"),
  header_default = quick_section_header("Landscape section"),
  footer_default = quick_section_footer("Landscape section")
))
doc <- rtf_add(doc, landscape_section)
doc <- rtf_add(doc, "Landscape orientation", style = "heading 2")
doc <- quick_hello_world(doc)

doc <- rtf_add(
  doc,
  block_section(
    prop_section(
      type = "nextPage",
      header_default = quick_section_header("Final section"),
      footer_default = quick_section_footer("Final section")
    )
  )
)
doc <- rtf_add(doc, "Back to portrait", style = "heading 2")
doc <- quick_hello_world(doc)
print(doc, target = tempfile(fileext = ".rtf"))
#> [1] "/tmp/RtmpZCXWQh/file18db3501e861.rtf"
```
