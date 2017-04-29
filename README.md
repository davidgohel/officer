officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->
The officer package lets R users manipulate Word (`.docx`) and PowerPoint (`*.pptx`) documents. In short, one can add images, tables and text into documents from R.

*This package is close to ReporteRs as it produces Word and PowerPoint files but it is faster, do not require `rJava` (but `xml2`) and has less functions that will make it easier to maintain.*

[![Build Status](https://travis-ci.org/davidgohel/officer.svg?branch=master)](https://travis-ci.org/davidgohel/officer) [![CRAN version](http://www.r-pkg.org/badges/version/officer)](https://CRAN.R-project.org/package=officer) [![Coverage Status](https://img.shields.io/codecov/c/github/davidgohel/officer/master.svg)](https://codecov.io/github/davidgohel/officer?branch=master) [![Project Status: WIP - Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](http://www.repostatus.org/badges/latest/wip.svg)](http://www.repostatus.org/#wip) ![cranlogs](http://cranlogs.r-pkg.org./badges/officer)

Word documents
--------------

Function `read_docx` will read an initial Word document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into a Word document:

-   images: produce your plot in png or emf files and add them into the document; as a whole paragraph or inside a paragraph.
-   tables: add data.frames as tables, format is defined by the associated Word table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined by the associated Word paragraph and text styles.
-   field codes: add Word field codes inside paragraphs. Field codes is an old feature of MS Word to create calculated elements such as tables of content, automatic numbering and hyperlinks.

In a Word document, one can use cursor functions to reach the beginning of a document, its end or a particular paragraph containing a given text. This *cursor* concept has been implemented to make easier the post processing of files.

The file generation is performed with function `print`.

### import Word document in a data.frame

Function `docx_summary` read and import content of a Word document into a tibble object. The function handles paragraphs, tables and section breaks.

PowerPoint documents
--------------------

Function `read_pptx` will read an initial PowerPoint document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into existing or new PowerPoint slides:

-   images: produce your plot in png or emf files and add them in a slide.
-   tables: add data.frames as tables, format is defined by the associated PowerPoint table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined in the corresponding layout of the slide.

In a PowerPoint document, one can set a slide as selected and reach a particular shape (and remove it or add text).

The file generation is performed with function `print`.

### import PowerPoint document in a data.frame

Function `pptx_summary` read and import content of a PowerPoint document into a tibble object. The function handles paragraphs, tables and images.

### Tables and package `flextable`

The package [flextable](https://github.com/davidgohel/flextable) brings a full API to produce nice tables and use them with `officer`.

Installation
------------

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
```

Or the latest version on CRAN:

``` r
install.packages("officer")
```
