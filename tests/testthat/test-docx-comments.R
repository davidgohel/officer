test_that("add comments", {
  fp_bold <- fp_text_lite(bold = TRUE)
  fp_red <- fp_text_lite(color = "red")

  bl <- block_list(
    fpar(ftext("Comment multiple words.", fp_bold)),
    fpar(
      ftext("Second line.", fp_red)
    )
  )

  a_par <- fpar(
    "This paragraph contains",
    run_comment(
      cmt = bl,
      run = ftext("a comment."),
      author = "Author Me",
      date = "2023-06-01"
    )
  )

  b_par <- fpar(
    run_comment(
      cmt = block_list(
        fpar(
          ftext("Comment a paragraph.")
        )
      ),
      run = ftext("This paragraph contains another comment."),
      author = "Author You",
      date = "2023-06-01",
      initials = "OM"
    )
  )
  doc <- read_docx()
  doc <- body_add_fpar(doc, value = a_par, style = "Normal")
  doc <- body_add_fpar(doc, value = b_par, style = "Normal")

  docx_file <- print(doc, target = tempfile(fileext = ".docx"))
  docx_dir <- tempfile()
  unpack_folder(docx_file, docx_dir)

  doc <- read_xml(file.path(docx_dir, "word/comments.xml"))
  comment1 <- xml_find_first(doc, "w:comment[@w:id='0']")
  comment2 <- xml_find_first(doc, "w:comment[@w:id='1']")

  expect_false(inherits(comment1, "xml_missing"))
  expect_false(inherits(comment2, "xml_missing"))

  expect_length(xml_children(comment1), 2)
  expect_length(xml_children(comment2), 1)
})

test_that("docx_comments accounts for multiple comments in a paragraph", {
  multi_comment_par <- fpar(
    "This paragraph",
    run_comment(
      cmt = block_list(
        fpar(
          ftext("First Comment.")
        )
      ),
      run = ftext("contains"),
      author = "Author Me",
      date = "2023-06-01"
    ),
    "multiple",
    run_comment(
      cmt = block_list(
        fpar(
          ftext("Second Comment.")
        )
      ),
      run = ftext("comments"),
      author = "Author Me",
      date = "2023-06-01"
    )
  )

  doc <- read_docx()
  doc <- body_add_fpar(doc, value = multi_comment_par, style = "Normal")
  docx_file <- print(doc, target = tempfile(fileext = ".docx"))

  comments <- docx_comments(read_docx(docx_file))

  expect_equal(nrow(comments), 2)

  expect_equal(
    comments$text,
    list("First Comment.", "Second Comment.")
  )
  expect_equal(
    comments$commented_text,
    list("contains", "comments")
  )
})

test_that("docx_comments", {
  example_docx <- "docs_dir/test-docx_comments.docx"
  doc <- read_docx(path = example_docx)

  comments <- docx_comments(doc)

  # Doc includes 16 comments
  expect_equal(nrow(comments), 16)
  # No NAs in "commented_text"
  expect_true(
    all(
      vapply(
        comments[["commented_text"]],
        function(x) all(!is.na(x)),
        FUN.VALUE = logical(1)
      )
    )
  )
  # Accounts for empty comments or multi line comments
  expect_true(
    all(
      lengths(comments[["text"]][-c(1, 4)]) == 1
    )
  )
  ## Comment 1 has 2 lines
  expect_equal(
    length(comments[["text"]][[1]]),
    2
  )
  ## Comment 4 is empty
  expect_identical(
    comments[["text"]][[4]],
    character(0)
  )
})

test_that("docx_comments accounts for comments spanning no or multiple paragraphs", {
  example_docx <- "docs_dir/test-docx_comments.docx"
  doc <- read_docx(path = example_docx)

  comments <- docx_comments(doc)

  expect_true(
    all(
      lengths(comments[["para_id"]][-c(5, 6)]) == 1
    )
  )
  ## Comment 5 spans no paragraph
  expect_identical(
    comments[["para_id"]][[5]],
    character(0)
  )
  expect_equal(
    paste(comments[["commented_text"]][[5]], collapse = " "),
    ""
  )
  ## Comment 6 spans 2 paragraphs
  expect_equal(
    length(comments[["para_id"]][[6]]),
    2
  )
  expect_equal(
    paste(comments[["commented_text"]][[6]], collapse = " "),
    "a comment … … which spans multiple paragraphs."
  )
})

test_that("docx_comments accounts for comments spanning no or multiple runs", {
  example_docx <- "docs_dir/test-docx_comments.docx"
  doc <- read_docx(path = example_docx)

  comments <- docx_comments(doc)

  expect_true(
    all(
      lengths(comments[["commented_text"]][-c(5, 6, 7, 15)]) == 1
    )
  )
  ## Comment 5 spans no run
  expect_identical(
    comments[["commented_text"]][[5]],
    character(0)
  )
  expect_equal(
    paste(comments[["commented_text"]][[5]], collapse = " "),
    ""
  )
  ## Comment 6 spans 2 runs as it spans 2 paragraphs
  expect_equal(
    length(comments[["commented_text"]][[6]]),
    2
  )
  expect_equal(
    paste(comments[["commented_text"]][[6]], collapse = " "),
    "a comment … … which spans multiple paragraphs."
  )
  ## Comment 7 spans 3 runs.
  expect_equal(
    length(comments[["commented_text"]][[7]]),
    3
  )
  expect_equal(
    paste(comments[["commented_text"]][[7]], collapse = ""),
    "spanning multiple runs."
  )
  ## Comment 15 spans 3 runs because of an inner comment
  expect_equal(
    length(comments[["commented_text"]][[15]]),
    3
  )
  expect_equal(
    paste(comments[["commented_text"]][[15]], collapse = ""),
    "This paragraph contains two nested comments."
  )
})

test_that("docx_comments accounts for replies", {
  example_docx <- "docs_dir/test-docx_comments.docx"
  doc <- read_docx(path = example_docx)

  comments <- docx_comments(doc)

  # Make "unique" id based on commented text and paragraph id
  comments$unique_id <- paste(
    comments$para_id,
    comments$commented_text,
    sep = "."
  )
  comments$unique_id <- factor(
    comments$unique_id,
    levels = unique(comments$unique_id)
  )
  comments$unique_id <- as.integer(comments$unique_id)

  comments_split <- split(
    comments,
    comments$unique_id
  )

  # Accouting for replies we have only 12 comments
  expect_equal(
    length(comments_split),
    12
  )
  expect_equal(
    vapply(
      comments_split[-c(8, 9, 10)],
      nrow,
      FUN.VALUE = integer(1),
      USE.NAMES = FALSE
    ),
    rep(1, 9)
  )
  ## 8 has two replies
  expect_equal(
    nrow(comments_split[[8]]),
    3
  )
  expect_equal(
    unique(unlist(comments[8:10, "commented_text"])),
    comments[["commented_text"]][[8]]
  )
  expect_equal(
    unique(unlist(comments[8:10, "para_id"])),
    comments[["para_id"]][[8]]
  )
  ## 9 and 10 have one reply each
  expect_equal(
    nrow(comments_split[[9]]),
    2
  )
  expect_equal(
    unique(unlist(comments[11:12, "commented_text"])),
    comments[["commented_text"]][[11]]
  )
  expect_equal(
    unique(unlist(comments[11:12, "para_id"])),
    comments[["para_id"]][[11]]
  )
  expect_equal(
    nrow(comments_split[[10]]),
    2
  )
  expect_equal(
    unique(unlist(comments[13:14, "commented_text"])),
    comments[["commented_text"]][[13]]
  )
  expect_equal(
    unique(unlist(comments[13:14, "para_id"])),
    comments[["para_id"]][[13]]
  )
})

test_that("docx_comments accounts for nested comments", {
  example_docx <- "docs_dir/test-docx_comments.docx"
  doc <- read_docx(path = example_docx)

  comments <- docx_comments(doc)

  ## Outer Comment 15 spans 3 runs
  expect_equal(
    length(comments[["commented_text"]][[15]]),
    3
  )
  expect_equal(
    paste(comments[["commented_text"]][[15]], collapse = ""),
    "This paragraph contains two nested comments."
  )
  expect_equal(
    paste(comments[["text"]][[15]], collapse = ""),
    "Outer Comment."
  )

  ## Inner Comment 16 spans 1 run
  expect_equal(
    length(comments[["commented_text"]][[16]]),
    1
  )
  expect_equal(
    paste(comments[["commented_text"]][[16]], collapse = ""),
    "contains two "
  )
  expect_equal(
    paste(comments[["text"]][[16]], collapse = ""),
    "Inner Comment."
  )
})
