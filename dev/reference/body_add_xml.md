# Add an xml string as document element

Add an xml string as document element in the document. This function is
to be used to add custom openxml code.

## Usage

``` r
body_add_xml(x, str, pos = c("after", "before", "on"))
```

## Arguments

- x:

  an rdocx object

- str:

  a wml string

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".
