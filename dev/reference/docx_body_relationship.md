# Body xml document

Get the body document as xml. This function is not to be used by end
users, it has been implemented to allow other packages to work with
officer.

## Usage

``` r
docx_body_relationship(x)
```

## Arguments

- x:

  an rdocx object

## Examples

``` r
doc <- read_docx()
docx_body_relationship(doc)
#> <relationship>
#>   Public:
#>     add: function (id, type, target, target_mode = NA) 
#>     add_drawing: function (src, root_target) 
#>     add_img: function (src, root_target) 
#>     clone: function (deep = FALSE) 
#>     feed_from_xml: function (path) 
#>     get_data: function () 
#>     get_images_path: function () 
#>     get_next_id: function () 
#>     initialize: function (id = character(0), type = character(0), target = character(0)) 
#>     remove: function (target) 
#>     show: function () 
#>     write: function (path) 
#>   Private:
#>     ext_src:      
#>     get_int_id: function () 
#>     id: rId3 rId2 rId1 rId6 rId5 rId4
#>     target: settings.xml styles.xml numbering.xml theme/theme1.xml f ...
#>     target_mode: NA NA NA NA NA NA
#>     type: http://schemas.openxmlformats.org/officeDocument/2006/re ...
```
