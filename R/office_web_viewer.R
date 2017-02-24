#' @export
#' @title Office Web Viewer
#' @description Produce an iframe linked to Office Web Viewer. It let's you
#' display a Microsoft Office document in a Web browser.
#' @param url file url
#' @param width iframe width
#' @param height iframe height
office_web_viewer <- function(url, width = "80%", height="500px"){
  stopifnot(requireNamespace("htmltools", quietly = TRUE))
  htmltools::tags$iframe(
    "iframes are not supported",
    src = paste0("https://view.officeapps.live.com/op/view.aspx?src=", url),
    width = width, height=height
  )
}

