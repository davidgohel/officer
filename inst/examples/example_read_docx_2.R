# Using a custom template ----
# This example shows how to start from an existing template
# instead of creating a blank document

# Get the path to a landscape template included in the package
template <- system.file(package = "officer", "doc_examples", "landscape.docx")

# Create a document based on the template
# The document will inherit the template's styles and page settings
doc_2 <- read_docx(path = template)

# Add a section with a table
doc_2 <- body_add_par(doc_2, "Motor Trend Car Data", style = "heading 2")
doc_2 <- body_add_table(doc_2, value = head(mtcars))

# Add a section with a plot
doc_2 <- body_add_par(doc_2, "Sales Distribution", style = "heading 2")
doc_2 <- body_add_plot(doc_2, distribution_plot)

# Save the document
docx_file_output <- print(doc_2, target = tempfile(fileext = ".docx"))
