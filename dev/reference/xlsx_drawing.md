# Drawing manager for xlsx sheets

R6 class that manages a drawing file for an xlsx sheet. Used internally
by
[`sheet_add_drawing()`](https://davidgohel.github.io/officer/dev/reference/sheet_add_drawing.md)
methods.

## Super class

`officer::openxml_document` -\> `xlsx_drawing`

## Methods

### Public methods

- [`xlsx_drawing$new()`](#method-xlsx_drawing-new)

- [`xlsx_drawing$add_chart_anchor()`](#method-xlsx_drawing-add_chart_anchor)

- [`xlsx_drawing$add_chart_rel()`](#method-xlsx_drawing-add_chart_rel)

- [`xlsx_drawing$add_image_rel()`](#method-xlsx_drawing-add_image_rel)

- [`xlsx_drawing$add_image_anchor()`](#method-xlsx_drawing-add_image_anchor)

- [`xlsx_drawing$clone()`](#method-xlsx_drawing-clone)

Inherited methods

- [`officer::openxml_document$dir_name()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-dir_name)
- [`officer::openxml_document$feed()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-feed)
- [`officer::openxml_document$file_name()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-file_name)
- [`officer::openxml_document$get()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-get)
- [`officer::openxml_document$name()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-name)
- [`officer::openxml_document$rel_df()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-rel_df)
- [`officer::openxml_document$relationship()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-relationship)
- [`officer::openxml_document$remove()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-remove)
- [`officer::openxml_document$replace_xml()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-replace_xml)
- [`officer::openxml_document$save()`](https://davidgohel.github.io/officer/dev/reference/openxml_document.html#method-save)

------------------------------------------------------------------------

### Method `new()`

Create or reuse a drawing for a sheet.

#### Usage

    xlsx_drawing$new(package_dir, sheet_obj, content_type)

#### Arguments

- `package_dir`:

  path to the unpacked xlsx directory

- `sheet_obj`:

  a sheet R6 object

- `content_type`:

  a content_type R6 object

------------------------------------------------------------------------

### Method `add_chart_anchor()`

Add a chart anchor to the drawing.

#### Usage

    xlsx_drawing$add_chart_anchor(
      chart_rid,
      from_col = 3L,
      from_row = 1L,
      to_col = 10L,
      to_row = 15L
    )

#### Arguments

- `chart_rid`:

  relationship id of the chart

- `from_col`:

  top-left column anchor (0-based)

- `from_row`:

  top-left row anchor (0-based)

- `to_col`:

  bottom-right column anchor (0-based)

- `to_row`:

  bottom-right row anchor (0-based)

------------------------------------------------------------------------

### Method `add_chart_rel()`

Add a relationship from the drawing to a chart file.

#### Usage

    xlsx_drawing$add_chart_rel(chart_basename)

#### Arguments

- `chart_basename`:

  filename of the chart XML

------------------------------------------------------------------------

### Method `add_image_rel()`

Add a relationship from the drawing to a media image.

#### Usage

    xlsx_drawing$add_image_rel(image_basename)

#### Arguments

- `image_basename`:

  filename (without directory) of the image sitting in `xl/media/`.

------------------------------------------------------------------------

### Method `add_image_anchor()`

Add an image anchor to the drawing (absolute placement in inches from
the top-left corner of the sheet).

#### Usage

    xlsx_drawing$add_image_anchor(
      image_rid,
      left = 1,
      top = 1,
      width = 2,
      height = 2,
      alt = ""
    )

#### Arguments

- `image_rid`:

  relationship id of the image

- `left, top`:

  top-left anchor in inches

- `width, height`:

  size in inches

- `alt`:

  alternative text

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    xlsx_drawing$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
