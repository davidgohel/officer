x <- read_pptx()

layout_summary(x) # available layouts

x <- add_slide(x, layout = "Two Content")

x <- layout_default(x, "Title Slide") # set default layout for `add_slide()`
x <- add_slide(x) # uses default layout

# use `...` to fill placeholders when adding slide
x <- add_slide(
  x,
  layout = "Two Content",
  `Title 1` = "A title",
  dt = "Jan. 26, 2025",
  `body[2]` = "Body 2",
  left = "Left side",
  `6` = "Footer"
)
