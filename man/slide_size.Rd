% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/pptx_informations.R
\name{slide_size}
\alias{slide_size}
\title{Slides width and height}
\usage{
slide_size(x)
}
\arguments{
\item{x}{an rpptx object}
}
\description{
Get the width and height of slides in inches as
a named vector.
}
\examples{
my_pres <- read_pptx()
my_pres <- add_slide(my_pres,
  layout = "Two Content", master = "Office Theme")
slide_size(my_pres)
}
\seealso{
Other functions for reading presentation information: 
\code{\link{annotate_base}()},
\code{\link{color_scheme}()},
\code{\link{doc_properties}()},
\code{\link{layout_properties}()},
\code{\link{layout_summary}()},
\code{\link{length.rpptx}()},
\code{\link{plot_layout_properties}()},
\code{\link{slide_summary}()}
}
\concept{functions for reading presentation information}
