# Drawing manager for xlsx sheets

R6 class that manages a drawing file for an xlsx sheet. Used internally
by
[`sheet_add_drawing()`](https://davidgohel.github.io/officer/reference/sheet_add_drawing.md)
methods.

## Super class

`openxml_document` -\> `xlsx_drawing`

## Methods

### Public methods

- [`xlsx_drawing$new()`](#method-xlsx_drawing-initialize)

- [`xlsx_drawing$add_chart_anchor()`](#method-xlsx_drawing-add_chart_anchor)

- [`xlsx_drawing$add_chart_rel()`](#method-xlsx_drawing-add_chart_rel)

- [`xlsx_drawing$add_image_rel()`](#method-xlsx_drawing-add_image_rel)

- [`xlsx_drawing$add_image_anchor()`](#method-xlsx_drawing-add_image_anchor)

- [`xlsx_drawing$clone()`](#method-xlsx_drawing-clone)

Inherited methods

- `openxml_document$dir_name()`
- `openxml_document$feed()`
- `openxml_document$file_name()`
- `openxml_document$get()`
- `openxml_document$name()`
- `openxml_document$rel_df()`
- `openxml_document$relationship()`
- `openxml_document$remove()`
- `openxml_document$replace_xml()`
- `openxml_document$save()`

------------------------------------------------------------------------

### `xlsx_drawing$new()`

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

### `xlsx_drawing$add_chart_anchor()`

Place a chart on the sheet. The chart can be positioned in three ways,
matching Excel's own "Format Picture" options:

- Move and size with cells (Excel's default): supply `from` and `to` as
  cell references (e.g. `"B2"` and `"H20"`). The chart spans those two
  cells and follows them when rows/columns are inserted or resized.

- Move but don't size with cells: supply `from` only. The chart anchors
  to that cell with the size set by `width`/`height`.

- Fixed position: leave `from` and `to` `NULL`. The chart is placed at
  `left`/`top` with size `width`/`height`, independent of cell layout.

#### Usage

    xlsx_drawing$add_chart_anchor(
      chart_rid,
      left = 1,
      top = 1,
      width = 6,
      height = 4,
      from = NULL,
      to = NULL,
      edit_as = c("twoCell", "oneCell", "absolute"),
      graphic_uri = ooxml_chart_uris("drawingml")$graphic_uri,
      chart_inner = NULL
    )

#### Arguments

- `chart_rid`:

  relationship id of the chart (returned by `add_chart_rel()`). Ignored
  when `chart_inner` is supplied explicitly.

- `left, top`:

  top-left position in inches (fixed-position mode)

- `width, height`:

  size in inches

- `from, to`:

  Excel cell references like `"C4"` (1-based, case-insensitive)

- `edit_as`:

  how Excel should treat the drawing when rows or columns are resized:
  `"twoCell"` (move and size with cells, the default), `"oneCell"` (move
  only) or `"absolute"` (neither). Only meaningful when both `from` and
  `to` are supplied.

- `graphic_uri`:

  chart family URI; defaults to the classic chart family. Use
  [`ooxml_chart_uris()`](https://davidgohel.github.io/officer/reference/ooxml_chart_uris.md)
  to obtain the value for the chartEx family.

- `chart_inner`:

  advanced. XML fragment referencing the chart part. Leave `NULL` to let
  the method build the right reference from `graphic_uri`.

------------------------------------------------------------------------

### `xlsx_drawing$add_chart_rel()`

Register a chart part with the drawing. Returns the relationship id, to
be passed to `add_chart_anchor()`.

#### Usage

    xlsx_drawing$add_chart_rel(
      chart_basename,
      rel_type = ooxml_chart_uris("drawingml")$rel_type
    )

#### Arguments

- `chart_basename`:

  filename of the chart XML

- `rel_type`:

  chart family relationship type; defaults to the classic chart family.
  Use
  [`ooxml_chart_uris()`](https://davidgohel.github.io/officer/reference/ooxml_chart_uris.md)
  to obtain the value for the chartEx family.

------------------------------------------------------------------------

### `xlsx_drawing$add_image_rel()`

Add a relationship from the drawing to a media image.

#### Usage

    xlsx_drawing$add_image_rel(image_basename)

#### Arguments

- `image_basename`:

  filename (without directory) of the image sitting in `xl/media/`.

------------------------------------------------------------------------

### `xlsx_drawing$add_image_anchor()`

Place an image on the sheet. Positioning works the same way as
`add_chart_anchor()`: supply `from` and `to` to move and size with
cells, `from` only to move (but not resize) with cells, or leave both
`NULL` for a fixed position.

#### Usage

    xlsx_drawing$add_image_anchor(
      image_rid,
      left = 1,
      top = 1,
      width = 2,
      height = 2,
      from = NULL,
      to = NULL,
      edit_as = c("twoCell", "oneCell", "absolute"),
      alt = ""
    )

#### Arguments

- `image_rid`:

  relationship id of the image (returned by `add_image_rel()`)

- `left, top`:

  top-left position in inches (fixed-position mode)

- `width, height`:

  size in inches

- `from, to`:

  Excel cell references like `"C4"` (1-based, case-insensitive)

- `edit_as`:

  how Excel should treat the drawing when rows or columns are resized:
  `"twoCell"` (move and size with cells, the default), `"oneCell"` (move
  only) or `"absolute"` (neither). Only meaningful when both `from` and
  `to` are supplied.

- `alt`:

  alternative text

------------------------------------------------------------------------

### `xlsx_drawing$clone()`

The objects of this class are cloneable with this method.

#### Usage

    xlsx_drawing$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
