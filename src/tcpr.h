#include "border.h"
#include "Rcpp.h"
using namespace Rcpp;

class tcpr
{
public:
  tcpr(std::string, std::string,
       int, int, int, int,
       int, int, int, int,
       bool, std::string, std::string,
       IntegerVector, IntegerVector,
       IntegerVector, IntegerVector,
       CharacterVector, IntegerVector
  );
  std::string a_tag();
  std::string w_tag();
  std::string css();

private:
  std::string vertical_align;
  std::string text_direction;
  int mb, mt, ml, mr;
  int shading_r, shading_g, shading_b, shading_a;
  bool do_bgimg;
  std::string bgimg_rid;
  std::string bgimg_path;
  border b,t,l,r;
};
