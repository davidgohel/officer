# Read paragraph styles defined on an RTF document

Return the data.frame of styles currently registered on the document.
Useful for debugging or for checking which built-in styles are available
before referencing them via `style = ...`.

## Usage

``` r
rtf_styles_info(x)
```

## Arguments

- x:

  an rtf object created by
  [`rtf_doc()`](https://davidgohel.github.io/officer/dev/reference/rtf_doc.md).

## Value

a data.frame with one row per style.

## See also

[`rtf_set_paragraph_style()`](https://davidgohel.github.io/officer/dev/reference/rtf_set_paragraph_style.md),
[`styles_info()`](https://davidgohel.github.io/officer/dev/reference/styles_info.md)

## Examples

``` r
rtf_styles_info(rtf_doc())
#>    style_id style_name      type rtf_index based_on
#> 1    Normal     Normal paragraph         1     <NA>
#> 2 heading 1  Heading 1 paragraph         2   Normal
#> 3 heading 2  Heading 2 paragraph         3   Normal
#> 4 heading 3  Heading 3 paragraph         4   Normal
#>                                                                                            rtf
#> 1             \\sl240\\slmult1\\ql\\sb0\\sa0\\fi0\\li0\\ri0 %font:Arial%\\fs22%ftcolor:black% 
#> 2 \\sl240\\slmult1\\ql\\sb120\\sa240\\fi0\\li0\\ri0 \\outlinelevel0\\b\\fs36%ftcolor:#1F4E79% 
#> 3  \\sl240\\slmult1\\ql\\sb80\\sa160\\fi0\\li0\\ri0 \\outlinelevel1\\b\\fs28%ftcolor:#2E75B6% 
#> 4  \\sl240\\slmult1\\ql\\sb60\\sa120\\fi0\\li0\\ri0 \\outlinelevel2\\b\\fs24%ftcolor:#5B9BD5% 
```
