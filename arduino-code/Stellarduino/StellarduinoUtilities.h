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

#include <math.h>
#include <avr/pgmspace.h>
#include "Arduino.h"

#ifndef AlignmentStarStruct
#define AlignmentStarStruct

// Structure for storing alignment star data.
// TODO: Deprecate this in favour of the Star classes.
struct AlignmentStar {
  String name;
  float vmag;
  float ra;
  float dec;
};

#endif

#ifndef SimpleStarStruct
#define SimpleStarStruct

// Structure for storing star data.
// TODO: Deprecate this in favour of the Star classes.
struct SimpleStar {
  String name;
  float vmag;
  float ra;
  float dec;
  float alt;
  float az;
  float time;
};

#endif

#ifndef StellarduinoUtilities_h
#define StellarduinoUtilities_h

class StellarduinoUtilities
{
public:
  static String rad2hm(float rad, boolean highPrecision = false);
  static String rad2dm(float rad, boolean highPrecision = false);
  static String padding(String str, int length);
};

#endif
