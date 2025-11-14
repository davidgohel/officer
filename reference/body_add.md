# Add content into a Word document

This function adds objects into a Word document. Values are added as new
paragraphs or tables.

This function is experimental and will replace the `body_add_*`
functions later. For now it is only to be used for successive additions
and cannot be used in conjunction with the `body_add_*` functions.

## Usage

``` r
body_add(x, value, ...)

# S3 method for class 'character'
body_add(x, value, style = NULL, ...)

# S3 method for class 'numeric'
body_add(x, value, style = NULL, format_fun = formatC, ...)

# S3 method for class 'factor'
body_add(x, value, style = NULL, format_fun = as.character, ...)

# S3 method for class 'fpar'
body_add(x, value, style = NULL, ...)

# S3 method for class 'data.frame'
body_add(
  x,
  value,
  style = NULL,
  header = TRUE,
  tcf = table_conditional_formatting(),
  alignment = NULL,
  ...
)

# S3 method for class 'block_caption'
body_add(x, value, ...)

# S3 method for class 'block_list'
body_add(x, value, ...)

# S3 method for class 'block_toc'
body_add(x, value, ...)

# S3 method for class 'external_img'
body_add(x, value, style = "Normal", ...)

# S3 method for class 'run_pagebreak'
body_add(x, value, style = NULL, ...)

# S3 method for class 'run_columnbreak'
body_add(x, value, style = NULL, ...)

# S3 method for class 'gg'
body_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  scale = 1,
  unit = "in",
  ...
)

# S3 method for class 'plot_instr'
body_add(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  unit = "in",
  ...
)

# S3 method for class 'block_pour_docx'
body_add(x, value, ...)

# S3 method for class 'block_section'
body_add(x, value, ...)
```

## Arguments

- x:

  an rdocx object

- value:

  object to add in the document. Supported objects are vectors,
  data.frame, graphics, block of formatted paragraphs, unordered list of
  formatted paragraphs, pretty tables with package flextable,
  'Microsoft' charts with package mschart.

- ...:

  further arguments passed to or from other methods. When adding a
  `ggplot` object or
  [plot_instr](https://davidgohel.github.io/officer/reference/plot_instr.md),
  these arguments will be used by png function. See method signatures to
  see what arguments can be used.

- style:

  paragraph style name. These names are available with function
  [styles_info](https://davidgohel.github.io/officer/reference/styles_info.md)
  and are the names of the Word styles defined in the base document (see
  argument `path` from
  [read_docx](https://davidgohel.github.io/officer/reference/read_docx.md)).

- format_fun:

  a function to be used to format values.

- header:

  display header if TRUE

- tcf:

  conditional formatting settings defined by
  [`table_conditional_formatting()`](https://davidgohel.github.io/officer/reference/table_conditional_formatting.md)

- alignment:

  columns alignement, argument length must match with columns length,
  values must be "l" (left), "r" (right) or "c" (center).

- width, height:

  plot size in units expressed by the unit argument. Defaults to a width
  of 6 and a height of 5 "in"ches.

- res:

  resolution of the png image in ppi

- scale:

  Multiplicative scaling factor, same as in ggsave

- unit:

  One of the following units in which the width and height arguments are
  expressed: "in", "cm" or "mm".

## Methods (by class)

- `body_add(character)`: add a character vector.

- `body_add(numeric)`: add a numeric vector.

- `body_add(factor)`: add a factor vector.

- `body_add(fpar)`: add a
  [fpar](https://davidgohel.github.io/officer/reference/fpar.md) object.
  These objects enable the creation of formatted paragraphs made of
  formatted chunks of text.

- `body_add(data.frame)`: add a data.frame object with
  [`block_table()`](https://davidgohel.github.io/officer/reference/block_table.md).

- `body_add(block_caption)`: add a
  [block_caption](https://davidgohel.github.io/officer/reference/block_caption.md)
  object. These objects enable the creation of set of formatted
  paragraphs made of formatted chunks of text.

- `body_add(block_list)`: add a
  [block_list](https://davidgohel.github.io/officer/reference/block_list.md)
  object. Use this function to add a list of block elements (e.g.
  paragraphs, images, tables) into a Word document in a more efficient
  way than with usual `body_add_*` functions. This function will add
  several elements in a faster way because the cursor is not calculated
  for each iteraction over the elements, as a consequence the function
  only append elements at the end of the document and does not allow to
  insert elements at a specific position.

- `body_add(block_toc)`: add a table of content (a
  [block_toc](https://davidgohel.github.io/officer/reference/block_toc.md)
  object).

- `body_add(external_img)`: add an image (a
  [external_img](https://davidgohel.github.io/officer/reference/external_img.md)
  object).

- `body_add(run_pagebreak)`: add a
  [run_pagebreak](https://davidgohel.github.io/officer/reference/run_pagebreak.md)
  object.

- `body_add(run_columnbreak)`: add a
  [run_columnbreak](https://davidgohel.github.io/officer/reference/run_columnbreak.md)
  object.

- `body_add(gg)`: add a ggplot object.

- `body_add(plot_instr)`: add a base plot with a
  [plot_instr](https://davidgohel.github.io/officer/reference/plot_instr.md)
  object.

- `body_add(block_pour_docx)`: pour content of an external docx file
  with with a
  [block_pour_docx](https://davidgohel.github.io/officer/reference/block_pour_docx.md)
  object

- `body_add(block_section)`: ends a section with a
  [block_section](https://davidgohel.github.io/officer/reference/block_section.md)
  object

## Illustrations

![](figures/body_add_doc_1.png)

## Examples

``` r
# This example demonstrates the versatility of body_add()
# by showing how to add various content types to a Word document

# Create a new Word document
x <- read_docx()

# Add a heading for the table of contents section
x <- body_add(x, "Table of content", style = "heading 1")

# Insert an automatic table of contents
# This will list all headings in the document
x <- body_add(x, block_toc())

# Insert a page break to start content on a new page
x <- body_add(x, run_pagebreak())

# Add a main section heading
x <- body_add(x, "Iris Dataset Sample", style = "heading 1")

# Add a data.frame as a table
# The first 6 rows of the iris dataset will be displayed
x <- body_add(x, head(iris), style = "table_template")

# Add another section heading
x <- body_add(x, "Alphabetic List", style = "heading 1")

# Add a character vector as paragraphs
# Each letter will appear as a separate paragraph with "Normal" style
x <- body_add(x, letters, style = "Normal")

# Insert a continuous section break
# Content continues on the same page but section properties can change
x <- body_add(x, block_section(prop_section(type = "continuous")))

# Add a plot created with R base graphics
# plot_instr() captures the plotting code to generate the image
x <- body_add(x, plot_instr(code = barplot(1:5, col = 2:6)))

# Change to landscape orientation for the next section
# This creates a new section with landscape page layout
x <- body_add(
  x,
  block_section(prop_section(page_size = page_size(orient = "landscape")))
)

# Save the document to a file
output_file <- tempfile(fileext = ".docx")
print(x, target = output_file)
```
