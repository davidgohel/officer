#include <Rcpp.h>
#include "border.h"
using namespace Rcpp;



// [[Rcpp::export]]
std::string a_border(int r, int g, int b, int a, std::string type, double width) {
  border border_(r, g, b, a, type, width);
  return border_.a_tag();
}

