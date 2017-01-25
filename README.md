officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->
[![Build Status](https://travis-ci.org/davidgohel/officer.svg?branch=master)](https://travis-ci.org/davidgohel/officer) [![Coverage Status](https://img.shields.io/codecov/c/github/davidgohel/officer/master.svg)](https://codecov.io/github/davidgohel/officer?branch=master) [![CRAN version](http://www.r-pkg.org/badges/version/officer)](https://cran.r-project.org/package=officer) ![](http://cranlogs.r-pkg.org/badges/grand-total/officer) [![Project Status: WIP - Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](http://www.repostatus.org/badges/latest/wip.svg)](http://www.repostatus.org/#wip)

![](http://upload.wikimedia.org/wikipedia/commons/f/f7/Artillerie_garde_imperiale.jpg)

officer
-------

The officer package provides lets R users manipulate Word documents (`.docx`). This package is close to ReporteRs as it produces Word files (PowerPoint will be implemented later) but it is faster, do not require `rJava` (but `xml2`).

Last but not least, *cursor* concept has been implemented to make it easier the post processing of Word files. One can use the cursor to reach a particular paragraph containing a given text, or the beginning of the document and the end.

The package `flextable` brings in addition a full API to produce nice tables and use them with `officer`.

> There is no vignette yet, for now, please search the R help file (major functions are all starting with `docx_`).

### Installation

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
devtools::install_github("davidgohel/flextable")
```
