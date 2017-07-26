class border
{
public:
  border(int, int, int, int, std::string, double);
  std::string w_tag(std::string);
  std::string a_tag(std::string);
  std::string a_tag();
  std::string css(std::string);

private:
  int red;
  int green;
  int blue;
  int alpha;
  std::string type;
  double width;
};


