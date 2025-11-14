# Update bookmark of an autonumber run

This function lets recycling a object made by
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)
by changing the bookmark value. This is useful to avoid calling
[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)
several times because of many tables.

## Usage

``` r
set_autonum_bookmark(x, bkm = NULL)
```

## Arguments

- x:

  an object of class
  [`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)

- bkm:

  bookmark id to associate with autonumber run. Value can only be made
  of alpha numeric characters, ':', -' and '\_'.

## See also

[`run_autonum()`](https://davidgohel.github.io/officer/reference/run_autonum.md)

## Examples

``` r
z <- run_autonum(
  seq_id = "tab", pre_label = "Table ",
  bkm = "anytable"
)
set_autonum_bookmark(z, bkm = "anothertable")
#> $seq_id
#> [1] "tab"
#> 
#> $pre_label
#> [1] "Table "
#> 
#> $post_label
#> [1] ": "
#> 
#> $bookmark
#> [1] "anothertable"
#> 
#> $bookmark_all
#> [1] FALSE
#> 
#> $pr
#> NULL
#> 
#> $start_at
#> NULL
#> 
#> $tnd
#> [1] 0
#> 
#> $tns
#> [1] "-"
#> 
#> attr(,"class")
#> [1] "run_autonum" "run"        
```
