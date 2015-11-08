#include "StellarduinoUtils.h"

String StellarduinoUtils::rad2hm(float rad, boolean highPrecision) {
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

String StellarduinoUtils::rad2dm(float rad, boolean highPrecision) {
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

String StellarduinoUtils::padding(String str, int length) {
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
}


