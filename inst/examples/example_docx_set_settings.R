library(officer)

txt_lorem <- rep(
  "Purus lectus eros metus turpis mattis platea praesent sed. ",
  50
)
txt_lorem <- paste0(txt_lorem, collapse = "")

header_first <- block_list(fpar(ftext("text for first page header")))
header_even <- block_list(fpar(ftext("text for even page header")))
header_default <- block_list(fpar(ftext("text for default page header")))
footer_first <- block_list(fpar(ftext("text for first page footer")))
footer_even <- block_list(fpar(ftext("text for even page footer")))
footer_default <- block_list(fpar(ftext("text for default page footer")))

ps <- prop_section(
  header_default = header_default,
  footer_default = footer_default,
  header_first = header_first,
  footer_first = footer_first,
  header_even = header_even,
  footer_even = footer_even
)

x <- read_docx()

x <- docx_set_settings(
  x = x,
  zoom = 2,
  list_separator = ",",
  even_and_odd_headers = TRUE
)

for (i in 1:20) {
  x <- body_add_par(x, value = txt_lorem)
}
x <- body_set_default_section(
  x,
  value = ps
)
print(x, target = tempfile(fileext = ".docx"))
