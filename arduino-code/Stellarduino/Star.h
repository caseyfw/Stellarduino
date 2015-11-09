/**
 * Star.h
 *
 * Represents a celestial object with known celestial coordinates.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#ifndef Star_h
#define Star_h

#include "Arduino.h"

class Star
{
public:
  Star(String name, float ra, float dec);
private:
  String name;
  float ra;
  float dec;
};

#endif
