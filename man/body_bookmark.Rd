% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/docx_add.R
\name{body_bookmark}
\alias{body_bookmark}
\title{Add bookmark in a 'Word' document}
\usage{
body_bookmark(x, id)
}
\arguments{
\item{x}{an rdocx object}

\item{id}{bookmark name}
}
\description{
Add a bookmark at the cursor location. The bookmark
is added on the first run of text in the current paragraph.
}
\examples{

# cursor_bookmark ----

doc <- read_docx()
doc <- body_add_par(doc, "centered text", style = "centered")
doc <- body_bookmark(doc, "text_to_replace")
}
