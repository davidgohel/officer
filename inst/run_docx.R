library(magrittr)
library(officer)
library(flextable)
library(ionicons)

src = as_png(name = "calendar", fill = "#FFE64D", width = 15, height = 15)
myft <- flextable(head(iris)) %>%
  theme_tron() %>%
  display(i = ~ Sepal.Length < 5, Sepal.Length = fpar(external_img(src, width = 15/72, height = 15/72)) ) %>%
  autofit()

src = as_png(name = "usb", fill = "red", width = 72, height = 72)

# doc <- docx() %>%
#   cursor_begin() %>%
#   docx_add_toc(pos = "on") %>%
#   docx_add_pbreak() %>%
#   docx_add_img(src, width = 1, height = 1, style = "Normal" ) %>%
#   # docx_add_par("Titre de niveau 1", style = "heading 1") %>%
#   # docx_add_par("Titre de niveau 2", style = "heading 2") %>%
#   docx_add_flextable( value = myft )

doc <- docx() %>%
  docx_add_toc(style = "centered") %>%
  docx_add_flextable(value = myft) %>%
  docx_add_par("", style = "Normal") %>%
  docx_add_par("", style = "Normal") %>%
  docx_append_run("sympa", style = "strong") %>%
  docx_append_img(src, width = .2, height = .2, style = "strong") %>%
  docx_append_seqfield(str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT", style = 'Default Paragraph Font') %>%
  docx_add_img(src, width = 1, height = 1, style = "Normal" ) %>%
  docx_add_par(" : Liste des machin", style = "centered") %>%
  docx_append_seqfield(str = "SEQ Figure \u005C* roman", style = 'Default Paragraph Font', pos = "before") %>%
  docx_append_run("Figure ", style = "strong", pos = "before") %>%
  docx_add_img(src, width = 1, height = 1, style = "Normal" ) %>%
  docx_add_par(" : Liste 2 de machins", style = "centered") %>%
  docx_append_seqfield(str = "SEQ Figure \u005C* roman", style = 'Default Paragraph Font', pos = "before")  %>%
  docx_append_run("Figure ", style = "strong", pos = "before") %>%
  docx_add_par("", style = "Normal") %>%
  docx_append_seqfield(str = "SYMBOL 100 \u005Cf Wingdings", style = 'Default Paragraph Font') %>%
  docx_add_table(iris, style = "table_template", width = 3, last_column = TRUE)


unlink(c("aaaa", "test3.docx" ), recursive = TRUE, force = TRUE)

doc <- docx() %>%
  docx_append_img(src = as_png(name = "calendar", fill = "#FFE64D", width = 15, height = 15),
                  width = .2, height = .2, style = "strong")
print(doc, target = "test3.docx")
unzip("test3.docx", exdir = "aaaa")
#
browseURL("test3.docx")
