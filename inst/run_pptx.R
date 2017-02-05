library(tidyverse)
library(officer)
library(flextable)

myft <- flextable(head(iris)) %>%
  theme_vanilla() %>%
  autofit()


unlink(c("aaaa", "coco.pptx" ), recursive = TRUE, force = TRUE)
doc <- pptx(path = "example.pptx")
doc <- add_slide(doc, layout = "Titre et contenu", master = "masque1")
doc <- set_title(x = doc, value = "Titre")
doc <- add_slide(doc, layout = "Titre et contenu", master = "masque2") %>%
  pptx_add_flextable(value = myft, offx = 2/3, offy = 1.75)
print(doc, target = "coco.pptx" ) %>% browseURL()

