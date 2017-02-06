library(tidyverse)
library(officer)
library(flextable)

myft <- flextable(head(iris)) %>%
  theme_vanilla() %>%
  autofit()


unlink(c("aaaa", "coco.pptx" ), recursive = TRUE, force = TRUE)
doc <- pptx(path = "example.pptx")
doc <- add_slide(doc, layout = "Titre et contenu", master = "masque1")
doc <- set_title(x = doc, value = "Titre coco")
doc <- add_slide(doc, layout = "Titre et contenu", master = "masque2")
print(doc, target = "coco.pptx" ) %>% browseURL()
doc$slideLayouts$get_master()$get_data()
# slide_info(doc, layout = "Titre et contenu", master = "masque1")
