officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->

> Make corporate reporting with minimum hassle

[![R build
status](https://github.com/davidgohel/officer/workflows/R-CMD-check/badge.svg)](https://github.com/davidgohel/officer/actions)
[![version](https://www.r-pkg.org/badges/version/officer)](https://CRAN.R-project.org/package=officer)
[![codecov test
coverage](https://codecov.io/gh/davidgohel/officer/branch/master/graph/badge.svg)](https://app.codecov.io/gh/davidgohel/officer)

The officer package lets R users manipulate Word (`.docx`) and
PowerPoint (`*.pptx`) documents. In short, one can add images, tables
and text into documents from R. An initial document can be provided;
contents, styles and properties of the original document will then be
available. It also supports the writing of ‘RTF’ documents and basic
manipulation of Excel (‘.xlsx’) workbooks.

## Ressources

The help pages are in a bookdown located at:

<https://ardata-fr.github.io/officeverse/>

Manuals are available at:

<https://davidgohel.github.io/officer/>.

## Word documents

The `read_docx()` function will read an initial Word document (an empty
one by default) and lets you modify its content later.

The package provides functions to add R outputs into a Word document:

- images: produce your plot in png or emf files and add them into the
  document, as a whole paragraph or inside a paragraph.
- tables: add data.frames as tables, format is defined by the associated
  Word table style.
- text: add text as paragraphs or inside an existing paragraph, format
  is defined by the associated Word paragraph and text styles.
- field codes: add Word field codes inside paragraphs. Field codes is an
  old feature of MS Word to create calculated elements such as tables of
  contents, automatic numbering and hyperlinks.

File generation is performed with the `print` function.

### import Word document in a data.frame

The function `docx_summary()` reads and imports content of a Word
document into a data.frame. The function handles paragraphs, tables and
section breaks. The function `docx_comments()` reads comments of a Word
document and organise the results into a data.frame.

## PowerPoint documents

The function `read_pptx()` will read an initial PowerPoint document (an
empty one by default) and let you modify its content later.

The package provides functions to add R outputs into existing or new
PowerPoint slides:

- images: produce your plot in png or emf files and add them in a slide.
- tables: add data.frames as tables, format is defined by the associated
  PowerPoint table style.
- text: add text as paragraphs or inside an existing paragraph, format
  is defined in the corresponding layout of the slide.

In a PowerPoint document, one can set a slide as selected and reach a
particular shape (and remove it or add text).

File generation is performed with the `print()` function.

### import PowerPoint document in a data.frame

The `pptx_summary()` function reads and imports content of a PowerPoint
document into a data.frame. The function handles paragraphs, tables and
images.

## Extensions

### Tables and package `flextable`

The package
[flextable](https://ardata-fr.github.io/flextable-book/index.html)
brings a full API to produce nice tables and use them with packages
officer and rmarkdown.

### Vector graphics with package `rvg`

The package [rvg](https://github.com/davidgohel/rvg) brings an API to
produce nice vector graphics that can be embedded in PowerPoint
documents or Excel workbooks with `officer`.

### Native office charts with package `mschart`

The package [mschart](https://github.com/ardata-fr/mschart) combined
with `officer` can produce native office charts in PowerPoint and Word
documents.

### Advance Word documents with R Markdown with package `officedown`

The package [officedown](https://ardata-fr.github.io/officeverse/)
facilitates the formatting of Microsoft Word documents produced by R
Markdown documents.

## Installation

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
```

Or the latest version on CRAN:

``` r
install.packages("officer")
```

## Getting help

If you have questions about how to use the package, visit Stack
Overflow’s `officer` tag and post your question there. I usually read
them and answer when possible.

## Contributing to the package

### Code of Conduct

Anyone getting involved in this package agrees to our [Code of
Conduct](https://github.com/davidgohel/officer/blob/master/CONDUCT.md).

### Bug reports

Bugs are reported through [GitHub
issues](https://github.com/davidgohel/officer/issues/new/choose). The
issue form requires a reproducible example (a
[reprex](https://reprex.tidyverse.org/)) and the output of
`sessionInfo()`; reports missing these elements are closed without
further response. The more reproducible your report, the more time can
go into investigating the bug rather than reconstructing it.

For usage questions (“how do I do X with officer”), please read
<https://ardata-fr.github.io/officeverse/> or
<https://davidgohel.github.io/officer/> or use [Stack
Overflow](https://stackoverflow.com/questions/tagged/officer) with the
`officer` tag instead.

### Pull requests

Before writing any code, please open an issue describing the change you
have in mind and the reason for it. The proposal needs to be discussed
and explicitly validated by the maintainer before work starts; pull
requests for features that have not gone through this step will not be
considered, no matter how well crafted. This isn’t meant to discourage
contributions, only to make sure the time you invest goes into a change
that will actually land.

This policy keeps the package coherent: officer’s internals are shared
with several downstream packages (flextable, officedown, rvg), and
changes that look local often have wider implications. Aligning on the
design in an issue first saves time on both sides.

Once a change has been agreed, the PR should include:

- the new function(s) with code and roxygen tags, including runnable
  examples
- documentation updates (vignette section if relevant)
- corresponding tests under `tests/testthat`

Run `rhub::check_for_cran()` locally before submitting; CI will run the
full check suite on the PR.
