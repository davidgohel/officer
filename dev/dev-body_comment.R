x <- read_docx()
ft <- fp_text(font.size = 12, bold = TRUE)
comment <- ftext("Hello Comment", ft)

x <- body_comment(x, comment)

print(x, "dev/write-comment-1.docx")
