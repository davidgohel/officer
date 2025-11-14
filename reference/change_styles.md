# Replace styles in a 'Word' Document

Replace styles with others in a 'Word' document. This function can be
used for paragraph, run/character and table styles.

## Usage

``` r
change_styles(x, mapstyles)
```

## Arguments

- x:

  an rdocx object

- mapstyles:

  a named list, names are the replacement style, content (as a character
  vector) are the styles to be replaced. Use
  [`styles_info()`](https://davidgohel.github.io/officer/reference/styles_info.md)
  to display available styles.

## Examples

``` r
# creating a sample docx so that we can illustrate how
# to change styles
doc_1 <- read_docx()

doc_1 <- body_add_par(doc_1, "A title", style = "heading 1")
doc_1 <- body_add_par(doc_1, "Another title", style = "heading 2")
doc_1 <- body_add_par(doc_1, "Hello world!", style = "Normal")
file <- print(doc_1, target = tempfile(fileext = ".docx"))

# now we can illustrate how
# to change styles with `change_styles`
doc_2 <- read_docx(path = file)
mapstyles <- list(
  "centered" = c("Normal", "heading 2"),
  "strong" = "Default Paragraph Font"
)
doc_2 <- change_styles(doc_2, mapstyles = mapstyles)
print(doc_2, target = tempfile(fileext = ".docx"))
```
