# Style manager for xlsx workbooks

R6 class that manages fonts, fills, borders, and cell formats in an xlsx
workbook's `styles.xml`. Used internally by
[`sheet_write_data()`](https://davidgohel.github.io/officer/reference/sheet_write_data.md)
and available for extensions.

## Methods

### Public methods

- [`xlsx_styles$new()`](#method-xlsx_styles-new)

- [`xlsx_styles$get_font_id()`](#method-xlsx_styles-get_font_id)

- [`xlsx_styles$get_fill_id()`](#method-xlsx_styles-get_fill_id)

- [`xlsx_styles$get_border_id()`](#method-xlsx_styles-get_border_id)

- [`xlsx_styles$get_style_id()`](#method-xlsx_styles-get_style_id)

- [`xlsx_styles$get_xf_id()`](#method-xlsx_styles-get_xf_id)

- [`xlsx_styles$save()`](#method-xlsx_styles-save)

- [`xlsx_styles$clone()`](#method-xlsx_styles-clone)

------------------------------------------------------------------------

### Method `new()`

Initialize styles from an xlsx package directory.

#### Usage

    xlsx_styles$new(package_dir)

#### Arguments

- `package_dir`:

  path to the unpacked xlsx directory

------------------------------------------------------------------------

### Method `get_font_id()`

Get or create a font index.

#### Usage

    xlsx_styles$get_font_id(
      name = "Calibri",
      size = 11,
      bold = FALSE,
      italic = FALSE,
      underline = FALSE,
      color = "000000"
    )

#### Arguments

- `name`:

  font family name

- `size`:

  font size in points

- `bold`:

  logical, bold

- `italic`:

  logical, italic

- `underline`:

  logical, underline

- `color`:

  hex color string (6 chars, without `#`)

------------------------------------------------------------------------

### Method `get_fill_id()`

Get or create a fill index.

#### Usage

    xlsx_styles$get_fill_id(bg_color = "FFFFFF")

#### Arguments

- `bg_color`:

  hex color string for fill background

------------------------------------------------------------------------

### Method `get_border_id()`

Get or create a border index.

#### Usage

    xlsx_styles$get_border_id(
      top_style = NULL,
      top_color = NULL,
      bottom_style = NULL,
      bottom_color = NULL,
      left_style = NULL,
      left_color = NULL,
      right_style = NULL,
      right_color = NULL
    )

#### Arguments

- `top_style`:

  border style for top side

- `top_color`:

  border color for top side

- `bottom_style`:

  border style for bottom side

- `bottom_color`:

  border color for bottom side

- `left_style`:

  border style for left side

- `left_color`:

  border color for left side

- `right_style`:

  border style for right side

- `right_color`:

  border color for right side

------------------------------------------------------------------------

### Method `get_style_id()`

Get or create a cell format (xf) index.

#### Usage

    xlsx_styles$get_style_id(
      font_id = 0L,
      fill_id = 0L,
      border_id = 0L,
      num_fmt_id = 0L,
      halign = NA,
      valign = NA,
      text_rotation = 0L,
      wrap_text = FALSE
    )

#### Arguments

- `font_id`:

  integer, font index (0-based)

- `fill_id`:

  integer, fill index (0-based)

- `border_id`:

  integer, border index (0-based)

- `num_fmt_id`:

  integer, number format id

- `halign`:

  horizontal alignment

- `valign`:

  vertical alignment

- `text_rotation`:

  text rotation angle (0-180)

- `wrap_text`:

  logical, enable text wrapping

------------------------------------------------------------------------

### Method `get_xf_id()`

Get or create a cell format index for a number format.

#### Usage

    xlsx_styles$get_xf_id(num_fmt_id)

#### Arguments

- `num_fmt_id`:

  integer, number format id

------------------------------------------------------------------------

### Method [`save()`](https://rdrr.io/r/base/save.html)

Save styles.xml to disk.

#### Usage

    xlsx_styles$save()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    xlsx_styles$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
