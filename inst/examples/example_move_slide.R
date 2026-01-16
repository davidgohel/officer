library(officer)

x <- read_pptx()
x <- add_slide(x, "Title and Content")
x <- ph_with(x, "Hello world 1", location = ph_location_type())
x <- add_slide(x, "Title and Content")
x <- ph_with(x, "Hello world 2", location = ph_location_type())
x <- move_slide(x, index = 1, to = 2)
