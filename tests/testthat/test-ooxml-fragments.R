test_that("solid_fill: builds a valid <a:solidFill> with srgbClr and alpha", {
  out <- solid_fill("red")
  expect_match(out, "^<a:solidFill>")
  expect_match(out, "</a:solidFill>$")
  expect_match(out, "<a:srgbClr val=\"FF0000\">")
  expect_match(out, "<a:alpha val=\"100000\"/>")
})

test_that("solid_fill: accepts standard '#hex' color strings", {
  expect_match(solid_fill("#3366CC"), "<a:srgbClr val=\"3366CC\">")
})

test_that("solid_fill: encodes alpha channel for semi-transparent colors", {
  half <- grDevices::rgb(1, 0, 0, alpha = 0.5)
  out <- solid_fill(half)
  # 0.5 round-trips through 8-bit (128/255) -> ~50196
  expect_match(out, "<a:alpha val=\"5[0-9]{4}\"/>")
})

test_that("solid_fill: 'transparent' produces alpha=0", {
  out <- solid_fill("transparent")
  expect_match(out, "<a:alpha val=\"0\"/>")
})

test_that("solid_fill: invalid color errors", {
  expect_error(solid_fill("not-a-color"))
})

test_that("to_pml.sp_line: emits <a:ln w=...> with solid fill and dash", {
  ln <- sp_line(color = "blue", lwd = 2, lty = "dash")
  out <- to_pml(ln)
  expect_match(out, "^<a:ln w=\"25400\"") # 2pt = 25400 EMU
  expect_match(out, "<a:srgbClr val=\"0000FF\">")
  expect_match(out, "<a:prstDash val=\"dash\"/>")
})

test_that("to_pml.sp_line: transparent or zero-width line emits <a:noFill/>", {
  ln <- sp_line(color = "transparent")
  out <- to_pml(ln)
  expect_match(out, "<a:noFill/>")
})

test_that("to_pml.sp_line: respects linecmpd, lineend, linejoin", {
  ln <- sp_line(
    color = "black",
    lwd = 1,
    linecmpd = "dbl",
    lineend = "sq",
    linejoin = "miter"
  )
  out <- to_pml(ln)
  expect_match(out, "cap=\"sq\"")
  expect_match(out, "cmpd=\"dbl\"")
  expect_match(out, "<a:miter/>")
})

test_that("to_pml.sp_line: head/tail line ends are emitted", {
  ln <- sp_line(
    color = "black",
    lwd = 1,
    headend = sp_lineend(type = "arrow", width = "lg", length = "med"),
    tailend = sp_lineend(type = "triangle")
  )
  out <- to_pml(ln)
  expect_match(out, "<a:headEnd type=\"arrow\" w=\"lg\" len=\"med\"/>")
  expect_match(out, "<a:tailEnd type=\"triangle\"")
})
