x <- read_docx()
ft <- fp_text(font.size = 12, bold = TRUE)
comment <- run_comment(ftext("Hello Comment", ft))
paragraph <- fpar("Hello World, ", comment, "Hello World, ")

x <- body_add_fpar(x, paragraph)

print(x, "dev/write-comment-1.docx")

xx <- read_docx("dev/write-comment.docx")
