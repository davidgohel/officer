# Write data to a sheet

Write a content into a sheet of an xlsx workbook. Multiple calls can
write to different positions on the same sheet.

This is a generic function dispatching on `value`. Supported inputs:

- `data.frame`: written as a table (header in row `start_row`, data
  starting at `start_row + 1`).

- `character`: each element in its own cell. The `character` method
  accepts a `direction` argument (`"vertical"` – default – or
  `"horizontal"`) to stack elements in a column or a row.

- [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md):
  richtext paragraph written into a single cell at
  `(start_row, start_col)`. Font, size, colour, bold, italic, underline,
  strikethrough and sub/superscript chunks are honoured.

- [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md):
  one cell per `fpar` item. Also accepts `direction = "vertical"`
  (default) or `"horizontal"`.

## Usage

``` r
sheet_write_data(x, value, sheet, start_row = 1L, start_col = 1L, ...)
```

## Arguments

- x:

  rxlsx object

- value:

  a `data.frame`, `character` vector,
  [`fpar()`](https://davidgohel.github.io/officer/dev/reference/fpar.md)
  or
  [`block_list()`](https://davidgohel.github.io/officer/dev/reference/block_list.md)

- sheet:

  sheet name (must already exist)

- start_row:

  row index where the header / first cell will be written (default 1)

- start_col:

  column index where the first column / first cell will be written
  (default 1)

- ...:

  method-specific arguments. In particular, `direction = "vertical"`
  (default) or `"horizontal"` is honoured by the `character` and
  `block_list` methods.

## Examples

``` r
# sheet_write_data() is an S3 generic dispatching on `value`:
#   - data.frame
#   - character
#   - fpar
#   - block_list

library(officer)

wb <- read_xlsx()
# drop the template's default sheet so the workbook contains only `demo`
wb <- sheet_remove(wb, sheet = sheet_names(wb)[1])
wb <- add_sheet(wb, label = "demo")

# --- A1:C4  data.frame (tabular write) ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  value = head(iris, 3)[, 1:3],
  start_row = 1,
  start_col = 1
)

# --- A6:A8  character, vertical default ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  value = c("first", "second", "third"),
  start_row = 6,
  start_col = 1
)

# --- A10:C10  character, horizontal ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  value = c("one", "two", "three"),
  direction = "horizontal",
  start_row = 10,
  start_col = 1
)

# --- A12  fpar with rich text in a single cell ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  start_row = 12,
  start_col = 1,
  value = fpar(
    ftext("bold ", fp_text(bold = TRUE, color = "red", font.size = 12)),
    ftext("italic ", fp_text(italic = TRUE, color = "blue")),
    ftext("under ", fp_text(underlined = TRUE)),
    ftext("strike ", fp_text(strike = TRUE, color = "#888888")),
    ftext("H", fp_text()),
    ftext("2", fp_text(vertical.align = "subscript")),
    ftext("O  (m", fp_text()),
    ftext("2", fp_text(vertical.align = "superscript")),
    ftext(")", fp_text())
  )
)

# --- A14:A16  block_list of fpars, stacked vertically ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  start_row = 14,
  start_col = 1,
  value = block_list(
    fpar(ftext(
      "header row",
      fp_text(bold = TRUE, font.size = 14, color = "#006699")
    )),
    fpar(ftext("middle row", fp_text(italic = TRUE, color = "#222222"))),
    fpar(ftext(
      "final row",
      fp_text(bold = TRUE, color = "darkgreen", font.size = 11)
    ))
  )
)

# --- B14:D14  block_list horizontal, 3 fpars in a row ----
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  start_row = 14,
  start_col = 2,
  direction = "horizontal",
  value = block_list(
    fpar(ftext("left", fp_text(bold = TRUE))),
    fpar(ftext("middle", fp_text(color = "red"))),
    fpar(ftext("right", fp_text(italic = TRUE)))
  )
)

# --- Commentary column (E) via character ---
wb <- sheet_write_data(
  wb,
  sheet = "demo",
  start_row = 1,
  start_col = 5,
  value = c(
    "data.frame -> tabular (iris cols 1-3)",
    "",
    "",
    "character(3) vertical -> A6:A8",
    "",
    "",
    "",
    "character(3) horizontal -> A10:C10",
    "",
    "",
    "fpar with 9 ftext chunks in A12",
    "",
    "",
    "block_list(3) vertical -> A14:A16",
    "",
    "",
    "(B14:D14) block_list horizontal"
  )
)

out <- tempfile(fileext = ".xlsx")
print(wb, target = out)
```
