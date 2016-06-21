class color_spec
{
public:
  color_spec (int, int, int, int);
  int is_visible();
  int has_alpha();
  std::string solid_fill();
  std::string solid_fill_w14();
  std::string w_color();
  std::string w_shd();
  std::string get_hex();
  std::string get_css();
  int is_transparent();

private:
  int red;
  int green;
  int blue;
  int alpha;
};
