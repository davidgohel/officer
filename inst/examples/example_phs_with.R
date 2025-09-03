library(officer)

# use key-value format to fill phs
x <- read_pptx()
x <- add_slide(x, "Two Content")
x <- phs_with(x,
  `Title 1` = "A title", # ph label
  dt = Sys.Date(), # ph type
  `body[2]` = "Body 2", # ph type + type index
  left = "Left side", # ph keyword
  `6` = "Footer" # ph index
)

# reuse ph content via the .dots arg
x <- read_pptx()
my_ph_list <- list(`6` = "Footer", dt = Sys.Date())
x <- add_slide(x, "Two Content")
x <- phs_with(x, `Title 1` = "Title A", `body[2]` = "Body A", .dots = my_ph_list)
x <- add_slide(x, "Two Content")
x <- phs_with(x, `Title 1` = "Title B", `body[2]` = "Body B", .dots = my_ph_list)

# use the .slide_idx arg to select which slide(s) to process
x <- read_pptx()
x <- add_slide(x, "Two Content")
x <- add_slide(x, "Two Content")
x <- phs_with(x, `6` = "Footer", dt = Sys.Date(), .slide_idx = 1:2)

# run to open temp pptx file locally
\dontrun{
print(x, preview = TRUE)
}
