# Table options in a 'knitr' context

Get options for table rendering in a 'knitr' context. It should not be
used by the end user, but its documentation should be read as a place
where table options are documented when 'knitr' is used.

The function is a utility to facilitate the retrieval of table options
supported by the 'flextable', 'officedown' and 'officer' packages.

These options should be set with `knitr::opts_chunk$set()`. The names
and expected values are listed in the following sections.

## Usage

``` r
opts_current_table()
```

## Value

a list

## knitr chunk options for table captions

|                                                  |                |           |
|--------------------------------------------------|----------------|-----------|
| **label**                                        | **name**       | **value** |
| caption id/bookmark                              | tab.id         | NULL      |
| caption                                          | tab.cap        | NULL      |
| display table caption on top of the table or not | tab.topcaption | TRUE      |
| caption table sequence identifier.               | tab.lp         | "tab:"    |

## knitr chunk options for Word table captions

|                                                         |                 |                           |
|---------------------------------------------------------|-----------------|---------------------------|
| **label**                                               | **name**        | **value**                 |
| Word stylename to use for table captions.               | tab.cap.style   | NULL                      |
| prefix for numbering chunk (default to "Table ").       | tab.cap.pre     | Table                     |
| suffix for numbering chunk (default to ": ").           | tab.cap.sep     | " :"                      |
| title number depth                                      | tab.cap.tnd     | 0                         |
| separator to use between title number and table number. | tab.cap.tns     | "-"                       |
| caption prefix formatting properties                    | tab.cap.fp_text | fp_text_lite(bold = TRUE) |

## knitr chunk options for Word tables

|                                                                |                     |           |
|----------------------------------------------------------------|---------------------|-----------|
| **label**                                                      | **name**            | **value** |
| the Word stylename to use for tables                           | tab.style           | NULL      |
| autofit' or 'fixed' algorithm.                                 | tab.layout          | "autofit" |
| value of the preferred width of the table in percent (base 1). | tab.width           | 1         |
| Alternative title text                                         | tab.alt.title       | NULL      |
| Alternative description text                                   | tab.alt.description | NULL      |

## knitr chunk options for data.frame with officedown

|                                                               |              |           |
|---------------------------------------------------------------|--------------|-----------|
| **label**                                                     | **name**     | **value** |
| apply or remove formatting from the first row in the table    | first_row    | TRUE      |
| apply or remove formatting from the first column in the table | first_column | FALSE     |
| apply or remove formatting from the last row in the table     | last_row     | FALSE     |
| apply or remove formatting from the last column in the table  | last_column  | FALSE     |
| don't display odd and even rows                               | no_hband     | TRUE      |
| don't display odd and even columns                            | no_vband     | TRUE      |

## returned elements

- cap.style (default: NULL)

- cap.pre (default: "Table ")

- cap.sep (default: ":")

- cap.tnd (default: 0)

- cap.tns (default: "-")

- cap.fp_text (default: `fp_text_lite(bold = TRUE)`)

- id (default: NULL)

- cap (default: NULL)

- alt.title (default: NULL)

- alt.description (default: NULL)

- topcaption (default: TRUE)

- style (default: NULL)

- tab.lp (default: "tab:")

- table_layout (default: "autofit")

- table_width (default: 1)

- first_row (default: TRUE)

- first_column (default: FALSE)

- last_row (default: FALSE)

- last_column (default: FALSE)

- no_hband (default: TRUE)

- no_vband (default: TRUE)

## See also

Other functions for officer extensions:
[`fortify_location()`](https://davidgohel.github.io/officer/reference/fortify_location.md),
[`get_reference_value()`](https://davidgohel.github.io/officer/reference/get_reference_value.md),
[`shape_properties_tags()`](https://davidgohel.github.io/officer/reference/shape_properties_tags.md),
[`str_encode_to_rtf()`](https://davidgohel.github.io/officer/reference/str_encode_to_rtf.md),
[`to_html()`](https://davidgohel.github.io/officer/reference/to_html.md),
[`to_pml()`](https://davidgohel.github.io/officer/reference/to_pml.md),
[`to_rtf()`](https://davidgohel.github.io/officer/reference/to_rtf.md),
[`to_wml()`](https://davidgohel.github.io/officer/reference/to_wml.md),
[`wml_link_images()`](https://davidgohel.github.io/officer/reference/wml_link_images.md)
