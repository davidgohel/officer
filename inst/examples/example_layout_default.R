# set and remove the default layout
x <- read_pptx()
layout_default(x) # no defaults
x <- layout_default(x, "Title and Content") # set default
layout_default(x)
x <- add_slide(x) # new slide with default layout
x <- layout_default(x, NULL) # remove default
layout_default(x) # no defaults

# use when repeatedly adding slides with same layout
x <- read_pptx()
x <- layout_default(x, "Title and Content")
x <- add_slide(x, title = "1. Slide", body = "Some content")
x <- add_slide(x, title = "2. Slide", body = "Some more content")
x <- add_slide(x, title = "3. Slide", body = "Even more content")
