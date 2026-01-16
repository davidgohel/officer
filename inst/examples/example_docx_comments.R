library(officer)

bl <- block_list(
  fpar("Comment multiple words."),
  fpar("Second line")
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

doc <- read_docx()
doc <- body_add_fpar(doc, value = a_par, style = "Normal")

docx_file <- print(doc, target = tempfile(fileext = ".docx"))

docx_comments(read_docx(docx_file))
