#ifndef FILL_
#define FILL_

class styles; // Forward declaration because fill is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "patternFill.h"
#include "gradientFill.h"

class fill {

  public:

    patternFill  patternFill_;
    gradientFill gradientFill_;

    fill(); // Default constructor
    fill(rapidxml::xml_node<>* fill, styles* styles);
};

#endif
