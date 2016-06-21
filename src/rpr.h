class rpr
{
public:
  rpr (double, bool, bool, bool,
       int, int, int, int,
       int, int, int, int,
       std::string, std::string);
  std::string a_tag();
  std::string w_tag();
  std::string css();

  double size;
  bool italic;
  bool bold;
  bool underlined;

  std::string vertical_align;

  int col_font_r, col_font_g, col_font_b, col_font_a;
  int col_shading_r, col_shading_g, col_shading_b, col_shading_a;

  std::string fontname;
};
