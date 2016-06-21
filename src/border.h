class border
{
public:
  border(int, int, int, int, std::string, int);
  std::string w_tag(std::string);
  std::string a_tag(std::string);
  std::string css(std::string);

private:
  int red;
  int green;
  int blue;
  int alpha;
  std::string type;
  int width;
};


