% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/pptx_informations.R
\name{color_scheme}
\alias{color_scheme}
\title{Color scheme of a PowerPoint file}
\usage{
color_scheme(x)
}
\arguments{
\item{x}{an rpptx object}
}
\description{
Get the color scheme of a
'PowerPoint' master layout into a data.frame.
}
\examples{
x <- read_pptx()
color_scheme ( x = x )
}
\seealso{
Other functions for reading presentation information: 
\code{\link{annotate_base}()},
\code{\link{doc_properties}()},
\code{\link{layout_properties}()},
\code{\link{layout_summary}()},
\code{\link{length.rpptx}()},
\code{\link{plot_layout_properties}()},
\code{\link{slide_size}()},
\code{\link{slide_summary}()}
}
\concept{functions for reading presentation information}
