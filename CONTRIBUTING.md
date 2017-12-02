## Contributing to the package

### Code of Conduct

Anyone getting involved in this package agrees to our [Code of Conduct](https://github.com/davidgohel/officer/blob/master/CONDUCT.md).

### Bug reports

When you file a [bug report](https://github.com/davidgohel/officer/issues), please spend some time making it easy for me to follow and reproduce. The more time you spend on making the bug report coherent, the more time I can dedicate to investigate the bug as opposed to the bug report.

### Contributing to the package development

A great way to start is to contribute an example or improve the documentation.

If you want to submit a Pull Request to integrate functions of yours, please provide:

* the new function(s) with code and roxygen tags (with examples)
* a new section in the appropriate vignette that describes how to use the new function
* add corresponding tests in directory `tests/testthat`.

By using rhub (run `rhub::check_for_cran()`), you will see if everything is ok.
When submitted, the PR will be evaluated automatically on travis and appveyor and you will be able to see if something broke.
