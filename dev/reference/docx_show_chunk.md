# Show underlying text tag structure

Show the structure of text tags at the current cursor. This is most
useful when trying to troubleshoot search-and-replace functionality
using
[`body_replace_all_text()`](https://davidgohel.github.io/officer/dev/reference/body_replace_all_text.md).

## Usage

``` r
docx_show_chunk(x)
```

## Arguments

- x:

  a docx device

## See also

[`body_replace_all_text()`](https://davidgohel.github.io/officer/dev/reference/body_replace_all_text.md)

## Examples

``` r
doc <- read_docx()
doc <- body_add_par(doc, "Placeholder one")
doc <- body_add_par(doc, "Placeholder two")

# Show text chunk at cursor
docx_show_chunk(doc)  # Output is 'Placeholder two'
#> 1 text nodes found at this cursor. 
#>   <w:t>: 'Placeholder two'
```
