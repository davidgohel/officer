x <- read_pptx()

# select layout explicitly
plot_layout_properties(x = x, layout = "Title Slide", master = "Office Theme")
plot_layout_properties(x = x, layout = "Title Slide") # no master needed if layout name unique
plot_layout_properties(x = x, layout = 1) # use layout index instead of name

# plot current slide's layout (default if no layout is passed)
x <- read_pptx()
x <- add_slide(x, "Title Slide")
plot_layout_properties(x)

# change appearance: what to show, font size, legend etc.
plot_layout_properties(x, layout = "Two Content", title = FALSE, type = FALSE, id = FALSE)
plot_layout_properties(x, layout = 4, cex = c(labels = .8, id = .7, type = .7))
plot_layout_properties(x, 1, legend = TRUE)

