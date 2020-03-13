predict_png_file <- function(file){
  gsub("\\.(docx|pptx)$", "%03d.png", basename(file))
}
predict_png_rx <- function(file){
  gsub("\\.(docx|pptx)$", "[0-9]*.png", basename(file))
}

office_to_image <- function( url, class = NULL, dest = "." ){
  file <- url
  cmd_ <- sprintf(
    "/Applications/LibreOffice.app/Contents/MacOS/soffice --headless \"-env:UserInstallation=file:///tmp/LibreOffice_Conversion_${USER}\" --convert-to pdf:writer_pdf_Export --outdir %s %s",
    dirname(file), file )
  system(cmd_, wait = TRUE)

  pdf_file <- gsub("\\.(docx|pptx)$", ".pdf", file)
  png_file <- predict_png_file(file)
  screen_copies <- pdftools::pdf_convert(pdf = pdf_file, format = "png", verbose = FALSE, filenames = file.path(dest, png_file))
  unlink(pdf_file, force = TRUE)
  screen_copies
}



office_compare <- function(office_file, case_dir = NULL, basedir = getwd()){

  if(is.null(case_dir))
    case_dir <- gsub("\\.(docx|pptx)$", "", basename(office_file))

  refdir <- file.path(basedir, "tests/figs", case_dir )
  if(!dir.exists(refdir)){
    dir.create(refdir, showWarnings = FALSE, recursive = TRUE)
  }

  png_file <- file.path(refdir, predict_png_rx(office_file))
  existing_files <- list.files(path = refdir, pattern = predict_png_rx(office_file))
  if( length(existing_files) < 1 ){
    office_to_image(office_file, dest = refdir)
  }


  newdir <- file.path(tempdir(), case_dir )
  unlink(newdir, force = TRUE, recursive = TRUE)
  dir.create(newdir, showWarnings = FALSE, recursive = TRUE)
  office_to_image(office_file, dest = newdir)

  z <- list()
  for(f in list.files(newdir, full.names = TRUE)){
    # browser()
    base_doc <- file.path(refdir, basename(f))
    outpng <- tempfile(fileext = ".png")
    x <- system(command = sprintf("pixelmatch %s %s %s 0.1", base_doc, f, outpng), intern = TRUE)
    x <- gsub("error: ", "", tail(x, n = 1), fixed = TRUE)
    x <- gsub("%", "", x)
    z[[basename(f)]] <- as.double(x)
  }
  z
}

zz <- list()
zz[[1 ]] <- office_compare("docs/articles/offcran/assets/docx/body_add_demo.docx")
zz[[2 ]] <- office_compare("docs/articles/offcran/assets/docx/columns_landscape_section.docx")
zz[[3 ]] <- office_compare("docs/articles/offcran/assets/docx/columns_section.docx")
zz[[4 ]] <- office_compare("docs/articles/offcran/assets/docx/cursor.docx")
zz[[5 ]] <- office_compare("docs/articles/offcran/assets/docx/first_example.docx")
# zz[[6 ]] <- office_compare("docs/articles/offcran/assets/docx/init_doc.docx")
# zz[[7 ]] <- office_compare("docs/articles/offcran/assets/docx/ipsum_doc.docx")
# zz[[8 ]] <- office_compare("docs/articles/offcran/assets/docx/landscape_section.docx")
# zz[[9 ]] <- office_compare("docs/articles/offcran/assets/docx/replace_doc.docx")
# zz[[10]] <- office_compare("docs/articles/offcran/assets/docx/replace_template.docx")
# zz[[11]] <- office_compare("docs/articles/offcran/assets/docx/section.docx")
# zz[[12]] <- office_compare("docs/articles/offcran/assets/docx/slip_in_demo.docx")
# zz[[13]] <- office_compare("docs/articles/offcran/assets/docx/toc_and_captions.docx")

