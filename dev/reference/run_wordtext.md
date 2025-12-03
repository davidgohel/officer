# Word chunk of text with a style

Format a chunk of text associated with a 'Word' character style. The
style is defined with its unique identifer.

## Usage

``` r
run_wordtext(text, style_id = NULL)
```

## Arguments

- text:

  text value, a single character value

- style_id:

  'Word' unique style identifier associated with the style to use.

## See also

[`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md)

Other run functions for reporting:
[`external_img()`](https://davidgohel.github.io/officer/dev/reference/external_img.md),
[`floating_external_img()`](https://davidgohel.github.io/officer/dev/reference/floating_external_img.md),
[`ftext()`](https://davidgohel.github.io/officer/dev/reference/ftext.md),
[`hyperlink_ftext()`](https://davidgohel.github.io/officer/dev/reference/hyperlink_ftext.md),
[`run_autonum()`](https://davidgohel.github.io/officer/dev/reference/run_autonum.md),
[`run_bookmark()`](https://davidgohel.github.io/officer/dev/reference/run_bookmark.md),
[`run_columnbreak()`](https://davidgohel.github.io/officer/dev/reference/run_columnbreak.md),
[`run_comment()`](https://davidgohel.github.io/officer/dev/reference/run_comment.md),
[`run_footnote()`](https://davidgohel.github.io/officer/dev/reference/run_footnote.md),
[`run_footnoteref()`](https://davidgohel.github.io/officer/dev/reference/run_footnoteref.md),
[`run_linebreak()`](https://davidgohel.github.io/officer/dev/reference/run_linebreak.md),
[`run_pagebreak()`](https://davidgohel.github.io/officer/dev/reference/run_pagebreak.md),
[`run_reference()`](https://davidgohel.github.io/officer/dev/reference/run_reference.md),
[`run_tab()`](https://davidgohel.github.io/officer/dev/reference/run_tab.md),
[`run_word_field()`](https://davidgohel.github.io/officer/dev/reference/run_word_field.md)

## Examples

``` r
run1 <- run_wordtext("hello", "DefaultParagraphFont")
paragraph <- fpar(run1)

x <- read_docx()
x <- body_add_fpar(x, paragraph)
print(x, target = tempfile(fileext = ".docx"))
```
