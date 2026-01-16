# Set 'Microsoft Word' Document Settings

Set various settings of a 'Microsoft Word' document generated with
'officer'. Options include:

- zoom factor (default view in Word),

- default tab stop,

- hyphenation zone,

- decimal symbol,

- list separator (see details below),

- compatibility mode,

- even and odd headers management (see details below),

- and auto hyphenation activation.

## Usage

``` r
docx_set_settings(
  x,
  zoom = 1,
  default_tab_stop = 0.5,
  hyphenation_zone = 0.25,
  decimal_symbol = ".",
  list_separator = ";",
  compatibility_mode = "15",
  even_and_odd_headers = FALSE,
  auto_hyphenation = FALSE,
  unit = "in"
)
```

## Arguments

- x:

  an rdocx object

- zoom:

  zoom factor, default is 1 (100%)

- default_tab_stop:

  default tab stop in inches, default is 0.5

- hyphenation_zone:

  hyphenation zone in inches, default is 0.25

- decimal_symbol:

  decimal symbol, default is "."

- list_separator:

  list separator, default is ";". Sets the separator used by Word for
  lists (see details).

- compatibility_mode:

  compatibility mode, default is "15"

- even_and_odd_headers:

  whether to use different headers for even and odd pages, default is
  FALSE. Enables the "Different Odd and Even Pages" feature in
  'Microsoft Word'.

- auto_hyphenation:

  whether to enable auto hyphenation, default is FALSE.

- unit:

  unit for `default_tab_stop` and `hyphenation_zone`, one of "in", "cm",
  "mm".

## Details

- `even_and_odd_headers`: If TRUE, 'Microsoft Word' will use different
  headers for odd and even pages ("Different Odd & Even Pages" feature
  in Word). This is useful for professional documents or reports that
  require alternating page layouts.

- `list_separator`: Sets the character used by 'Microsoft Word' to
  separate items in lists (for example, when inserting tables or lists
  in Word). This parameter affects how 'Microsoft Word' handles data
  import/export (CSV, etc.) and can be adapted to language or local
  conventions (e.g., ";" for French, "," for English).

## See also

[`read_docx()`](https://davidgohel.github.io/officer/dev/reference/read_docx.md)

## Examples

``` r
library(officer)

txt_lorem <- rep(
  "Purus lectus eros metus turpis mattis platea praesent sed. ",
  50
)
txt_lorem <- paste0(txt_lorem, collapse = "")

header_first <- block_list(fpar(ftext("text for first page header")))
header_even <- block_list(fpar(ftext("text for even page header")))
header_default <- block_list(fpar(ftext("text for default page header")))
footer_first <- block_list(fpar(ftext("text for first page footer")))
footer_even <- block_list(fpar(ftext("text for even page footer")))
footer_default <- block_list(fpar(ftext("text for default page footer")))

ps <- prop_section(
  header_default = header_default,
  footer_default = footer_default,
  header_first = header_first,
  footer_first = footer_first,
  header_even = header_even,
  footer_even = footer_even
)

x <- read_docx()

x <- docx_set_settings(
  x = x,
  zoom = 2,
  list_separator = ",",
  even_and_odd_headers = TRUE
)

for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}
x <- body_set_default_section(
  x,
  value = ps
)
print(x, target = tempfile(fileext = ".docx"))
```
