#include <Rcpp.h>
#include "rpr.h"
using namespace Rcpp;
#include <iostream>

// [[Rcpp::export]]
SEXP rpr_new(Rcpp::List compounds) {

  Rcpp::XPtr<rpr> ptr( new rpr(Rcpp::as<double>(compounds["font.size"]),
                               Rcpp::as<bool>(compounds["italic"]),
                               Rcpp::as<bool>(compounds["bold"]),
                               Rcpp::as<bool>(compounds["underlined"]),
                               Rcpp::as<int>(compounds["col_font_r"]),
                               Rcpp::as<int>(compounds["col_font_g"]),
                               Rcpp::as<int>(compounds["col_font_b"]),
                               Rcpp::as<int>(compounds["col_font_a"]),
                               Rcpp::as<int>(compounds["col_shading_r"]),
                               Rcpp::as<int>(compounds["col_shading_g"]),
                               Rcpp::as<int>(compounds["col_shading_b"]),
                               Rcpp::as<int>(compounds["col_shading_a"]),
                               Rcpp::as<std::string>(compounds["font.family"]),
                               Rcpp::as<std::string>(compounds["vertical.align"])
                            ), true );
  return ptr;
}

// [[Rcpp::export]]
std::string rpr_w(SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );
  std::stringstream os;
  os << ptr->w_tag();
  return os.str();
}

// [[Rcpp::export]]
std::string rpr_p(SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );

  std::stringstream os;
  os << ptr->a_tag();
  return os.str();
}


// [[Rcpp::export]]
std::string rpr_css(SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );

  std::stringstream os;
  os << ptr->css();
  return os.str();
}


// [[Rcpp::export]]
std::string chunk_w(std::string value, SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );

  std::stringstream os;
  os << "<w:r>";
  os << ptr->w_tag();
  os << "<w:t xml:space=\"preserve\">";
  os << value;
  os << "</w:t></w:r>";
  return os.str();
}

// [[Rcpp::export]]
std::string chunk_p(std::string value, SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );

  std::stringstream os;
  os << "<a:r>";
  os << ptr->a_tag();
  os << "<a:t>";
  os << value;
  os << "</a:t></a:r>";
  return os.str();
}


// [[Rcpp::export]]
std::string chunk_html(std::string value, SEXP fp) {
  Rcpp::XPtr<rpr> ptr( fp );

  std::stringstream os;
  os << "<span style=\"";
  os << ptr->css();
  os << "\">";
  os << value;
  os << "</span>";
  return os.str();
}






