#ifndef STROKE_
#define STROKE_

class styles; // Forward declaration because font is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class stroke {

  public:

    Rcpp::String style_;
    color        color_;

    stroke(); // default constructor
    stroke(rapidxml::xml_node<>* stroke, styles* styles);
};

#endif
