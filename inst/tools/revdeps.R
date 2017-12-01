rev_dep_lib <- "/Users/davidgohel/Documents/revdep_library/"

install.packages("huxtable", lib = rev_dep_lib, dependencies = c("Depends", "Imports", "LinkingTo", "Suggests") )

officers <- unique( c("magrittr",
"lazyeval","purrr","tibble","R6","stringr",
"htmltools","htmlwidgets",
"gdtools", "rmarkdown", "knitr",
"xml2", "dplyr",
"testthat","data.table",
"cellranger",
"writexl" ) )
install.packages(officers, lib = rev_dep_lib,
                 dependencies = c("Depends", "Imports", "LinkingTo", "Suggests") )
install.packages(c("cowplot", "stringi", "weights") , lib = rev_dep_lib,
                 dependencies = c("Depends", "Imports", "LinkingTo", "Suggests") )
install.packages(c("officer") , lib = rev_dep_lib,
                 dependencies = c("Depends", "Imports", "LinkingTo", "Suggests") )


library(devtools)
revdep_check(libpath = rev_dep_lib, ignore = "flextable")
