library(tidyverse)
library(officer)
library(flextable)
library(ionicons)

myft <- flextable(head(iris)) %>%
  theme_vanilla() %>%
  autofit()

calendar_src = as_png(name = "calendar", fill = "#FFE64D", width = 144, height = 144)

doc <- read_pptx(path = system.file(package = "officer", "template/template_fr.pptx") )
doc <- add_slide(doc, layout = "Titre et contenu", master = "masque1")
doc <- placeholder_set_text(x = doc, id = "title", str = "Un titre")
doc <- placeholder_set_text(x = doc, id = "ftr", str = "pied de page")
doc <- placeholder_set_text(x = doc, id = "dt", str = format(Sys.Date()))
doc <- placeholder_set_text(x = doc, id = "sldNum", str = "slide 1")#
doc <- pptx_add_flextable(doc, value = myft, id = "body")

doc <- add_slide(doc, layout = "Titre et contenu", master = "masque1")
doc <- placeholder_set_img(doc, src = calendar_src, id = "body", width = 2, height = 2)

doc <- add_slide(doc, layout = "Diapositive de titre", master = "masque2")
doc <- placeholder_set_text(x = doc, id = "subTitle", str = "Un gros titre")
doc <- placeholder_set_text(x = doc, id = "ctrTitle", str = "Un titre")


print(doc, target = "coco.pptx" ) %>% browseURL()
