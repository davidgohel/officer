library(magrittr)
library(officer)

run_num <- run_autonum(seq_id = "tab", pre_label = "tab. ")
caption <- block_caption("iris table",
                         style = "Normal", id = "iris_table",
                         autonum = run_num )

doc <- read_docx() %>%
  body_add("A title", style = "heading 1") %>%
  body_add("Hello world!", style = "Normal") %>%
  body_add(caption) %>%
  body_add(iris, style = "table_template")

print(doc, target = tempfile(fileext = ".docx") )
