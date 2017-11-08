#include <Rcpp.h>
using namespace Rcpp;
#include <iostream>
#include <string>
#include <iostream>
#include <climits>
#include <fstream>      // std::ofstream

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



// [[Rcpp::export]]
String wml_table(DataFrame x, std::string style_id,
                 int first_row = true, int last_row = false,
                 int first_column = false, int last_column = false,
                 int no_hband = false, int no_vband = false, int header = false) {
  int nrow = x.nrows();
  int ncol = x.size();

  std::stringstream os;

  os << "<w:tbl" <<
    " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"" <<
      " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" <<
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" <<
          " xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">";

  os << "<w:tblPr>";
  os << "<w:tblStyle w:val=\"" << style_id << "\"/><w:tblW/>";
  os << "<w:tblLook w:firstRow=\"" << first_row <<
    "\" w:lastRow=\"" << last_row <<
      "\" w:firstColumn=\"" << first_column <<
        "\" w:lastColumn=\"" << last_column <<
          "\" w:noHBand=\"" << no_hband <<
            "\" w:noVBand=\"" << no_vband << "\"/>";
  os << "</w:tblPr>";

  if( header ){
    os << "<w:tr><w:trPr><w:tblHeader/></w:trPr>";
    CharacterVector names_ = x.names();
    for(int j = 0 ; j < ncol ; j++){
      os << "<w:tc><w:trPr/><w:p><w:r><w:t>" <<
        Rf_translateCharUTF8(names_[j]) <<
          "</w:t></w:r></w:p></w:tc>";
    }
    os << "</w:tr>";
  }

  for(int i = 0 ; i < nrow ; i++){
    os << "<w:tr>";
    for(int j = 0 ; j < ncol ; j++){
      CharacterVector tmp=x[j];
      os << "<w:tc><w:trPr/><w:p><w:r><w:t>";
      os << Rf_translateCharUTF8(tmp[i]);
      os << "</w:t></w:r></w:p></w:tc>";
    }
    os << "</w:tr>";
  }
  os << "</w:tbl>";



  return os.str();

}





std::string base26(unsigned long v)
{
  char const digits[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  size_t const base = sizeof(digits) - 1;
  char result[sizeof(unsigned long)*CHAR_BIT + 1];
  char* current = result + sizeof(result);
  *--current = '\0';

  while (v != 0) {
    v--;
    *--current = digits[v % base];
    v /= base;
  }
  return current;
}

// [[Rcpp::export]]
CharacterVector as_col_ref(IntegerVector x){
  int n = x.size();
  CharacterVector out(n);
  for(int i = 0 ; i < n ; i ++ ){
    out[i] = base26(x[i]);
  }
  return out;
}


// [[Rcpp::export]]
void xls_table_prepare(DataFrame x, CharacterVector col_types,
                          CharacterVector tags_start,
                          CharacterVector tags_end,
                          CharacterVector col_ref,
                          int start_at_row,
                          std::string xsl_file,
                          std::string xml_file,
                          LogicalVector to_xml) {
  int nrow = x.nrows();
  int ncol = x.size();
  std::ofstream xsl, xml;
  xsl.open (xsl_file.c_str(), std::ofstream::out | std::ofstream::app);
  xml.open (xml_file.c_str(), std::ofstream::out);

  xml << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
  xml << "<sheetData xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n";

  std::stringstream os_header;
  os_header << "<row r=\"" << start_at_row << "\">";
  CharacterVector names_ = x.names();
  for(int j = 0 ; j < ncol ; j++){
    os_header << "<c ";
    os_header << "r=\"" << col_ref[j] << start_at_row << "\" ";
    os_header << "t=\"inlineStr\">";
    os_header << "<is><t>" << Rf_translateCharUTF8(names_[j]) << "</t></is>";
    os_header << "</c>";
  }
  os_header << "</row>";
  if( to_xml[0] )
    xml << os_header.rdbuf();
  else xsl << os_header.rdbuf();

  for(int i = 0 ; i < nrow ; i++){
    std::stringstream os;
    os << "<row r=\"" << start_at_row + i + 1 << "\">";
    for(int j = 0 ; j < ncol ; j++){
      CharacterVector tmp=x[j];
      os << "<c ";
      os << "r=\"" << col_ref[j] << start_at_row + i + 1 << "\" ";
      os << "t=\"" << col_types[j] << "\">";
      os << tags_start[j] << tmp[i] << tags_end[j];
      os << "</c>";
    }
    os << "</row>";

    if( to_xml[i+1] )
      xml << os.rdbuf();
    else xsl << os.rdbuf();
  }

  xml << "\n</sheetData>";
  xsl << "\n    </sheetData>\n  </xsl:template>\n</xsl:stylesheet>";

  xml.close();
  xsl.close();
}


