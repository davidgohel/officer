#include <Rcpp.h>
using namespace Rcpp;
#include <iostream>
#include <string>
#include <iostream>
// This is a simple example of exporting a C++ function to R. You can
// source this function into an R session using the Rcpp::sourceCpp
// function (or via the Source button on the editor toolbar). Learn
// more about Rcpp at:
//
//   http://www.rcpp.org/
//   http://adv-r.had.co.nz/Rcpp.html
//   http://gallery.rcpp.org/
//

// [[Rcpp::export]]
String pml_table(DataFrame x, std::string style_id,
                 int col_width, int row_height,
                 int first_row = true, int last_row = false,
                 int first_column = false, int last_column = false, int header = true) {
  int nrow = x.nrows();
  int ncol = x.size();

  std::stringstream os;

  os << "<a:tbl>";
  os << "<a:tblPr";
  if( first_row ) os << " firstRow=\"1\"";
  if( last_row ) os << " lastRow=\"1\"";
  if( first_column ) os << " firstColumn=\"1\"";
  if( last_column ) os << " lastColumn=\"1\"";
  os << ">";

  os << "<a:tableStyleId>" << style_id << "</a:tableStyleId>";
  os << "</a:tblPr>";

  os << "<a:tblGrid>";
  for(int j = 0 ; j < ncol ; j++){
    os << "<a:gridCol w=\"" << col_width << "\"/>";
  }
  os << "</a:tblGrid>";

  if( header ){
    os << "<a:tr h=\"" << row_height << "\">";
    CharacterVector names_ = x.names();
    for(int j = 0 ; j < ncol ; j++){
      os << "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>";
      os << Rf_translateCharUTF8(names_[j]);
      os << "</a:t></a:r></a:p></a:txBody></a:tc>";
    }
    os << "</a:tr>";
  }

  for(int i = 0 ; i < nrow ; i++){
    os << "<a:tr h=\"" << row_height << "\">";
    for(int j = 0 ; j < ncol ; j++){
      CharacterVector tmp=x[j];
      os << "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>";
      os << Rf_translateCharUTF8(tmp[i]);
      os << "</a:t></a:r></a:p></a:txBody></a:tc>";
    }
    os << "</a:tr>";
  }
  os << "</a:tbl>";



  return os.str();

}

