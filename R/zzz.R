utils::globalVariables(c(".data"))

.onLoad <- function(libname, pkgname) {
  # init officer options
  options(officer.date_format = "%Y-%m-%d")
}
