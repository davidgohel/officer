% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/docx_add.R
\name{body_add_xml}
\alias{body_add_xml}
\title{Add an xml string as document element}
\usage{
body_add_xml(x, str, pos = c("after", "before", "on"))
}
\arguments{
\item{x}{an rdocx object}

\item{str}{a wml string}

\item{pos}{where to add the new element relative to the cursor,
one of "after", "before", "on".}
}
\description{
Add an xml string as document element in the document. This function
is to be used to add custom openxml code.
}
\keyword{internal}
