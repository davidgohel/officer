#include "border.h"
#include "Rcpp.h"
using namespace Rcpp;

class ppr
{
public:
  ppr (std::string,
       int, int, int, int,
       int, int, int, int,
       IntegerVector, IntegerVector,
       IntegerVector, IntegerVector,
       CharacterVector, IntegerVector
       );
  std::string a_tag();
  std::string w_tag();
  std::string css();

private:
  std::string text_align;
  int pb, pt, pl, pr;
  int shading_r, shading_g, shading_b, shading_a;
  border b,t,l,r;
};
