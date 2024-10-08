% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/pptx_informations.R
\name{layout_summary}
\alias{layout_summary}
\title{Presentation layouts summary}
\usage{
layout_summary(x)
}
\arguments{
\item{x}{an rpptx object}
}
\description{
Get information about slide layouts and
master layouts into a data.frame. This function returns
a data.frame containing all layout and master names.
}
\examples{
my_pres <- read_pptx()
layout_summary ( x = my_pres )
}
\seealso{
Other functions for reading presentation information: 
\code{\link{annotate_base}()},
\code{\link{color_scheme}()},
\code{\link{doc_properties}()},
\code{\link{layout_properties}()},
\code{\link{length.rpptx}()},
\code{\link{plot_layout_properties}()},
\code{\link{slide_size}()},
\code{\link{slide_summary}()}
}
\concept{functions for reading presentation information}
