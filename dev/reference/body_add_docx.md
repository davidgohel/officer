# Add an external docx in a 'Word' document

Add content of a docx into an rdocx object.

The function is using a 'Microsoft Word' feature: when the document will
be edited, the content of the file will be inserted in the main
document.

This feature is unlikely to work as expected if the resulting document
is edited by another software. You can use function
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)
to import the content as an alternative.

The file is added when the method
[`print()`](https://rdrr.io/r/base/print.html) that produces the final
Word file is called, so don't remove file defined with `src` before.

## Usage

``` r
body_add_docx(x, src, pos = "after")
```

## Arguments

- x:

  an rdocx object

- src:

  docx filename, the path of the file must not contain any '&' and the
  basename must not contain any space.

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

## See also

Other functions for adding content:
[`body_add_blocks()`](https://davidgohel.github.io/officer/dev/reference/body_add_blocks.md),
[`body_add_break()`](https://davidgohel.github.io/officer/dev/reference/body_add_break.md),
[`body_add_caption()`](https://davidgohel.github.io/officer/dev/reference/body_add_caption.md),
[`body_add_fpar()`](https://davidgohel.github.io/officer/dev/reference/body_add_fpar.md),
[`body_add_gg()`](https://davidgohel.github.io/officer/dev/reference/body_add_gg.md),
[`body_add_img()`](https://davidgohel.github.io/officer/dev/reference/body_add_img.md),
[`body_add_par()`](https://davidgohel.github.io/officer/dev/reference/body_add_par.md),
[`body_add_plot()`](https://davidgohel.github.io/officer/dev/reference/body_add_plot.md),
[`body_add_table()`](https://davidgohel.github.io/officer/dev/reference/body_add_table.md),
[`body_add_toc()`](https://davidgohel.github.io/officer/dev/reference/body_add_toc.md),
[`body_append_start_context()`](https://davidgohel.github.io/officer/dev/reference/body_append_context.md),
[`body_import_docx()`](https://davidgohel.github.io/officer/dev/reference/body_import_docx.md)

## Examples

``` r
file1 <- tempfile(fileext = ".docx")
file2 <- tempfile(fileext = ".docx")
file3 <- tempfile(fileext = ".docx")
x <- read_docx()
x <- body_add_par(x, "hello world 1", style = "Normal")
print(x, target = file1)

x <- read_docx()
x <- body_add_par(x, "hello world 2", style = "Normal")
print(x, target = file2)

x <- read_docx(path = file1)
x <- body_add_break(x)
x <- body_add_docx(x, src = file2)
print(x, target = file3)
```
