# URIs for chart relationships in xlsx drawings

Returns the pair of URIs needed by
[xlsx_drawing](https://davidgohel.github.io/officer/reference/xlsx_drawing.md)
to embed a chart into a sheet, keeping the relationship type and graphic
URI consistent so Excel opens the file without repair.

## Usage

``` r
ooxml_chart_uris(kind = c("drawingml", "chartex"))
```

## Arguments

- kind:

  chart family. `"drawingml"` covers classic charts (bar, line, pie,
  scatter, area, bubble, radar, stock). `"chartex"` covers the newer
  family (boxplot, funnel, histogram, pareto, sunburst, treemap,
  waterfall).

## Value

a named list with elements `rel_type` and `graphic_uri`.

## Examples

``` r
ooxml_chart_uris("drawingml")
#> $rel_type
#> [1] "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
#> 
#> $graphic_uri
#> [1] "http://schemas.openxmlformats.org/drawingml/2006/chart"
#> 
ooxml_chart_uris("chartex")
#> $rel_type
#> [1] "http://schemas.microsoft.com/office/2014/relationships/chartEx"
#> 
#> $graphic_uri
#> [1] "http://schemas.microsoft.com/office/drawing/2014/chartex"
#> 
```
