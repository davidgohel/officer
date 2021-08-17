getncheck <- function(x, str) {
  child_ <- xml_child(x, str)
  expect_false(inherits(child_, "xml_missing"))
  child_
}



test_that("seqfield add ", {
  x <- read_docx()
  x <- body_add_fpar(x,
    fpar(
      "Time is: ",
      run_word_field(field = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT")
    ),
    style = "Normal"
  )

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal(xml_text(child_), "TIME \\@ \"HH:mm:ss\" \\* MERGEFORMAT")

  z <- fpar(
    "Figure: ",
    run_word_field(field = "SEQ Figure \u005C* roman"),
    " - This is a figure title"
  )
  x <- body_add_fpar(x, z, style = "centered", )

  node <- x$doc_obj$get_at_cursor()
  expect_equal(xml_text(node), "Figure: SEQ Figure \\* roman - This is a figure title")
})



test_that("hyperlink add ", {
  href_ <- "https://github.com/davidgohel"
  x <- read_docx()

  zz <- fpar(
    "Here is a link: ",
    hyperlink_ftext(text = "the link", href = href_)
  )
  x <- body_add_fpar(x, zz, style = "Normal")
  print(x, target = tempfile(fileext = ".docx"))

  rel_df <- x$doc_obj$rel_df()
  expect_true("https%3A//github.com/davidgohel" %in% rel_df$target)
  expect_equal(rel_df[rel_df$target == "https%3A//github.com/davidgohel", ]$target_mode, "External")
  expect_match(rel_df[rel_df$target == "https%3A//github.com/davidgohel", ]$type, "^http://schemas(.*)hyperlink$")

  node <- x$doc_obj$get_at_cursor()
  child_ <- getncheck(node, "w:hyperlink/w:r/w:t")
  expect_equal(xml_text(child_), "the link")
})
