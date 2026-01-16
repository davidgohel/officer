library(officer)

x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- remove_slide(x)

# Remove multiple slides at once
x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
x <- add_slide(x, "Title and Content")
x <- remove_slide(x, index = c(2, 4))
pptx_file <- print(x, target = tempfile(fileext = ".pptx"))
pptx_file
