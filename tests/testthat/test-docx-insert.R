context("insert into paragraphs of docx")


getncheck <- function(x, str){
  child_ <- xml_child(x, str)
  expect_false( inherits(child_, "xml_missing") )
  child_
}



test_that("seqfield add ", {
  x <- read_docx() %>%
    body_add_par("Time is: ", style = "Normal") %>%
    slip_in_seqfield(
      str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT")

  node <- x$doc_obj$get_at_cursor()
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='begin']")
  getncheck(node, "w:r/w:fldChar[@w:fldCharType='end']")

  child_ <- getncheck(node, "w:r/w:instrText")
  expect_equal( xml_attr(child_, "dirty"), "true")
  expect_equal( xml_text(child_), "TIME \\@ \"HH:mm:ss\" \\* MERGEFORMAT" )

  x <- body_add_par(x, " - This is a figure title", style = "centered") %>%
    slip_in_seqfield(str = "SEQ Figure \u005C* roman",
                     style = 'Default Paragraph Font', pos = "before") %>%
    slip_in_text("Figure: ", style = "strong", pos = "before")
  node <- x$doc_obj$get_at_cursor()
  expect_equal( xml_text(node), "Figure: SEQ Figure \\* roman - This is a figure title" )
})

