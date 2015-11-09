/**
 * StellarduinoUtilities.h
 *
 * Some type defines and utility functions used by Stellarduino.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include "StellarduinoUtilities.h"

String StellarduinoUtilities::rad2hm(float rad, boolean highPrecision) {
  if (rad < 0) rad = rad + 2.0 * M_PI;
  float hours = rad * 24.0 / (2.0 * M_PI);
  float minutes = (hours - floor(hours)) * 60.0;

  if (highPrecision)
  {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + ":" + padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  } else {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + "." + (String)int(floor((minutes - floor(minutes)) * 10.0));
  }
}

String StellarduinoUtilities::rad2dm(float rad, boolean highPrecision) {
  float degs = abs(rad) * 360.0 / (2.0 * M_PI);
  float minutes = (degs - floor(degs)) * 60.0;
  String sign = "+";
  if (rad < 0) sign = "-";

  if (highPrecision)
  {
    return sign + padding((String)int(floor(degs)), 2) + "*" + padding((String)int(floor(minutes)), 2) + ":" + padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  }
  {
    return sign + padding((String)int(floor(degs)), 2) + "*" + padding((String)int(floor(minutes)), 2);
  }
}

String StellarduinoUtilities::padding(String str, int length) {
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
}


