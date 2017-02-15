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


doc <- read_docx() %>%
  body_add_toc(style = "centered") %>%
  body_add_flextable(value = myft) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  slip_in_text("sympa", style = "strong") %>%
  slip_in_img(src, width = .2, height = .2, style = "strong") %>%
  slip_in_seqfield(str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT", style = 'Default Paragraph Font') %>%
  body_add_img(src, width = 1, height = 1, style = "Normal" ) %>%
  body_add_par(" : Liste des machin", style = "centered") %>%
  slip_in_seqfield(str = "SEQ Figure \u005C* roman", style = 'Default Paragraph Font', pos = "before") %>%
  slip_in_text("Figure ", style = "strong", pos = "before") %>%
  body_add_img(src, width = 1, height = 1, style = "Normal" ) %>%
  body_add_par(" : Liste 2 de machins", style = "centered") %>%
  slip_in_seqfield(str = "SEQ Figure \u005C* roman", style = 'Default Paragraph Font', pos = "before")  %>%
  slip_in_text("Figure ", style = "strong", pos = "before") %>%
  body_add_par("", style = "Normal") %>%
  slip_in_seqfield(str = "SYMBOL 100 \u005Cf Wingdings", style = 'Default Paragraph Font') %>%
  body_add_table(iris, style = "table_template", width = 3, last_column = TRUE)


unlink(c("aaaa", "test3.docx" ), recursive = TRUE, force = TRUE)

# doc <- read_docx() %>%
#   slip_in_img(src = as_png(name = "calendar", fill = "#FFE64D", width = 15, height = 15),
#                   width = .2, height = .2, style = "strong")
print(doc, target = "test3.docx")
# unzip("test3.docx", exdir = "aaaa")
#
browseURL("test3.docx")
