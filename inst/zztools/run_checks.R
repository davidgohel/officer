unlink("checks", recursive = TRUE, force = TRUE)

devtools::build(path = "checks")
tools::check_packages_in_dir("checks",
  check_args = c("--as-cran", ""),
  reverse = list(repos = getOption("repos")["CRAN"]))

