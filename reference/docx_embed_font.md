# Embed Fonts in a Word Document

Copy TrueType or OpenType font files from the local system into a Word
document. The font data is stored inside the `.docx` archive so that the
document renders correctly when opened on a system where the font is not
installed.

When only `font_family` is provided (without `regular`), the function
looks up font file paths on the local system via
[`gdtools::sys_fonts()`](https://davidgohel.github.io/gdtools/reference/sys_fonts.html)
and copies them into the document. A message shows the equivalent
explicit call for reproducibility.

## Usage

``` r
docx_embed_font(
  x,
  font_family,
  regular = NULL,
  bold = NULL,
  italic = NULL,
  bold_italic = NULL
)
```

## Arguments

- x:

  an rdocx object

- font_family:

  font family name as it will appear in the document. When `regular` is
  not provided, this name is used to look up font files via
  [`gdtools::sys_fonts()`](https://davidgohel.github.io/gdtools/reference/sys_fonts.html).

- regular:

  path to the regular .ttf/.otf font file. If `NULL`, font files are
  detected automatically via gdtools.

- bold:

  path to the bold .ttf/.otf font file (optional)

- italic:

  path to the italic .ttf/.otf font file (optional)

- bold_italic:

  path to the bold-italic .ttf/.otf font file (optional)

## Value

the rdocx object with embedded fonts

## Font licensing

Embedding a font in a document redistributes it. You must ensure that
the font license permits embedding. Fonts under the SIL Open Font
License (e.g. Liberation, Google Fonts) generally allow it. Many
commercial fonts restrict or prohibit embedding. Check the license of
the font before using this function.

## See also

[`gdtools::sys_fonts()`](https://davidgohel.github.io/gdtools/reference/sys_fonts.html),
[`gdtools::register_gfont()`](https://davidgohel.github.io/gdtools/reference/register_gfont.html),
[`docx_set_settings()`](https://davidgohel.github.io/officer/reference/docx_set_settings.md)

## Examples

``` r
# automatic detection (requires gdtools)
gdtools::register_liberationsans()
#> [1] TRUE
doc <- read_docx()
doc <- docx_embed_font(doc, font_family = "Liberation Sans")
#> Embedding "Liberation Sans" with detected font files:
#> docx_embed_font(
#>   x,
#>   font_family = "Liberation Sans",
#>   regular = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
#>   bold = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
#>   italic = "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
#>   bold_italic = "/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf"
#> )

# # explicit paths (no gdtools needed)
# doc <- docx_embed_font(
#   doc,
#   font_family = "My Font",
#   regular = "path/to/font-regular.ttf",
#   bold = "path/to/font-bold.ttf"
# )
```
