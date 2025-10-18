library(officer)

# Example 1: Creating a multi-column section ----
# This example demonstrates how to create a section with two columns,
# which is useful for newsletters, brochures, or academic papers

# Create sample text for demonstration
sample_text <- paste(
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
  "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
)
# Repeat text to fill columns
long_text <- paste(rep(sample_text, 5), collapse = " ")

# Need to use the same page margins each time
# a prop_section is created
page_margins <- page_mar()

# Create a new document
doc_1 <- read_docx()

# Add a title for the two-column section
doc_1 <- body_add_par(doc_1, "Newsletter Article", style = "heading 1")
doc_1 <- body_end_block_section(
  doc_1,
  block_section(
    prop_section(
      type = "continuous",
      page_margins = page_margins
    )
  )
)

# Define section properties for two columns
# - Two columns of equal width (3 inches each)
# - 0.5 inch spacing between columns
# - Continuous section (no page break)
two_column_section <- prop_section(
  type = "continuous",
  page_margins = page_margins,
  section_columns = section_columns(
    widths = c(3, 3),
    space = 0.5,
    sep = FALSE
  )
)

# Add content that will flow across two columns
doc_1 <- body_add_par(doc_1, "Introduction", style = "heading 2")
doc_1 <- body_add_fpar(
  doc_1,
  value = fpar(ftext(long_text), run_columnbreak()),
  style = "Normal"
)
doc_1 <- body_add_par(doc_1, "Main Content", style = "heading 2")
doc_1 <- body_add_par(doc_1, value = long_text, style = "Normal")

# End the two-column section
# Content after this point will return to single column layout
doc_1 <- body_end_block_section(doc_1, block_section(two_column_section))

# Add content in normal single-column layout
doc_1 <- body_add_par(doc_1, "Conclusion (Single Column)", style = "heading 1")
doc_1 <- body_add_par(doc_1, value = sample_text, style = "Normal")
default_sect_properties <- prop_section(
  type = "continuous",
  page_margins = page_margins
)
doc_1 <- body_set_default_section(doc_1, default_sect_properties)

# Save the document
output_file_1 <- tempfile(fileext = ".docx")
print(doc_1, target = output_file_1)


# Example 2: Section with custom header and footer ----
# This example shows how to create a specific section with its own
# header (containing R logo) and footer (with date), while other
# sections have no header/footer

# Get the path to R logo
img_path <- file.path(R.home("doc"), "html", "logo.jpg")

# Create header content with R logo aligned to the right
section_header <- block_list(
  fpar(
    external_img(src = img_path, height = 0.4, width = 0.4),
    ftext(
      " | Special Report Section",
      fp_text_lite(font.size = 10, color = "#4472C4")
    ),
    fp_p = fp_par(text.align = "right", padding.bottom = 6)
  )
)

# Create footer content with current date centered
section_footer <- block_list(
  fpar(
    ftext("Generated on: ", fp_text_lite(font.size = 9)),
    run_word_field(
      field = "Date \\@ \"MMMM d, yyyy\"",
      prop = fp_text_lite(font.size = 9, italic = TRUE)
    ),
    fp_p = fp_par(
      text.align = "center",
      padding.top = 6,
      border.top = fp_border(color = "#CCCCCC", width = 1)
    )
  )
)

# Define section properties with header and footer
# This section will have landscape orientation and custom margins
section_with_hf <- prop_section(
  page_size = page_size(orient = "landscape", width = 11, height = 8.5),
  page_margins = page_mar(
    bottom = 1,
    top = 1.2,
    right = 1,
    left = 1,
    header = 0.5,
    footer = 0.5
  ),
  type = "continuous",
  header_default = section_header,
  footer_default = section_footer
)

# Create a new document
doc_2 <- read_docx()

# Section 1: Normal section without header/footer
doc_2 <- body_add_par(doc_2, "Introduction", style = "heading 1")
doc_2 <- body_add_par(
  doc_2,
  "This is the introduction section with standard formatting."
)
doc_2 <- body_add_par(doc_2, value = sample_text, style = "Normal")

# Section 2: Special section WITH header and footer
# Add content for the special section
doc_2 <- body_add_par(doc_2, "Data Analysis Section", style = "heading 1")
doc_2 <- body_add_par(
  doc_2,
  "This section contains detailed analysis and is displayed in landscape format."
)

# Add a wide table that benefits from landscape orientation
analysis_data <- data.frame(
  Category = c("Sales", "Marketing", "R&D", "Operations"),
  Q1 = c(150, 85, 120, 95),
  Q2 = c(165, 92, 125, 102),
  Q3 = c(178, 88, 135, 108),
  Q4 = c(195, 98, 140, 115),
  Total = c(688, 363, 520, 420),
  Growth = c("13%", "7%", "8%", "10%")
)
doc_2 <- body_add_table(doc_2, value = analysis_data, style = "table_template")

# Add a chart
doc_2 <- body_add_par(doc_2, "Quarterly Trends", style = "heading 2")
trend_plot <- plot_instr({
  quarters <- paste0("Q", 1:4)
  sales <- c(150, 165, 178, 195)
  marketing <- c(85, 92, 88, 98)

  barplot(
    rbind(sales, marketing),
    beside = TRUE,
    names.arg = quarters,
    col = c("#4472C4", "#ED7D31"),
    border = NA,
    main = "Sales vs Marketing Budget",
    ylab = "Amount ($K)",
    xlab = "Quarter",
    legend.text = c("Sales", "Marketing"),
    args.legend = list(x = "topleft")
  )
})
doc_2 <- body_add_plot(doc_2, trend_plot, width = 7, height = 4)

# End the section with header/footer
# This applies the landscape orientation and header/footer to this section only
doc_2 <- body_end_block_section(doc_2, block_section(section_with_hf))

# Section 3: Return to normal section without header/footer
doc_2 <- body_add_par(doc_2, "Conclusions", style = "heading 1")
doc_2 <- body_add_par(
  doc_2,
  "This final section returns to portrait orientation without header or footer."
)
doc_2 <- body_add_par(doc_2, value = sample_text, style = "Normal")

# Save the document
output_file_2 <- tempfile(fileext = ".docx")
print(doc_2, target = output_file_2)
