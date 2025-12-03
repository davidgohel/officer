# Ensure valid slide indexes

Ensure valid slide indexes

## Usage

``` r
stop_if_not_in_slide_range(x, idx, arg = NULL, call = parent.frame())
```

## Arguments

- x:

  An `rpptx` object.

- idx:

  Slide indexes.

- arg:

  Name of argument to use in error message (optional).

- call:

  Environment to display in error message. Defaults to caller env. Set
  `NULL` to suppress (see
  [cli::cli_abort](https://cli.r-lib.org/reference/cli_abort.html)).
