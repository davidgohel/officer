getncheck <- function(x, str) {
  child_ <- xml_child(x, str)
  expect_false(inherits(child_, "xml_missing"))
  child_
}


test_that("seqfield add ", {
  x <- read_docx()
  x <- body_add_fpar(
    x,
    fpar(
      "Time is: ",
      run_word_field(field = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT")
    ),
    style = "Normal"
  )

  node <- docx_current_block_xml(x)
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

  node <- docx_current_block_xml(x)
  expect_equal(
    xml_text(node),
    "Figure: SEQ Figure \\* roman - This is a figure title"
  )
})
