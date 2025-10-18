# This example demonstrates the versatility of body_add()
# by showing how to add various content types to a Word document

# Create a new Word document
x <- read_docx()

# Add a heading for the table of contents section
x <- body_add(x, "Table of content", style = "heading 1")

# Insert an automatic table of contents
# This will list all headings in the document
x <- body_add(x, block_toc())

# Insert a page break to start content on a new page
x <- body_add(x, run_pagebreak())

# Add a main section heading
x <- body_add(x, "Iris Dataset Sample", style = "heading 1")

# Add a data.frame as a table
# The first 6 rows of the iris dataset will be displayed
x <- body_add(x, head(iris), style = "table_template")

# Add another section heading
x <- body_add(x, "Alphabetic List", style = "heading 1")

# Add a character vector as paragraphs
# Each letter will appear as a separate paragraph with "Normal" style
x <- body_add(x, letters, style = "Normal")

# Insert a continuous section break
# Content continues on the same page but section properties can change
x <- body_add(x, block_section(prop_section(type = "continuous")))

# Add a plot created with R base graphics
# plot_instr() captures the plotting code to generate the image
x <- body_add(x, plot_instr(code = barplot(1:5, col = 2:6)))

# Change to landscape orientation for the next section
# This creates a new section with landscape page layout
x <- body_add(
  x,
  block_section(prop_section(page_size = page_size(orient = "landscape")))
)

# Save the document to a file
output_file <- tempfile(fileext = ".docx")
print(x, target = output_file)
