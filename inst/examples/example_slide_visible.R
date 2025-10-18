path <- system.file("doc_examples/example.pptx", package = "officer")
x <- read_pptx(path)

slide_visible(x) # get slide visibilities

x <- slide_visible(x, hide = 1:2) # hide slides 1 and 2
x <- slide_visible(x, show = 1:2) # make slides 1 and 2 visible
x <- slide_visible(x, show = 1:2, hide = 3)

slide_visible(x) <- FALSE # hide all slides
slide_visible(x) <- c(TRUE, FALSE, TRUE) # set each slide separately
slide_visible(x) <- c(TRUE, FALSE) # warns that rhs values are recycled

slide_visible(x)[2] <- TRUE # set 2nd slide to visible
slide_visible(x)[c(1, 3)] <- FALSE # 1st and 3rd slide
slide_visible(x)[c(1, 3)] <- c(FALSE, FALSE) # identical
