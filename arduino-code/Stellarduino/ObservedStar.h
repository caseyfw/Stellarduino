/**
 * ObservedStar.h
 *
 * Extends Star class to represent a celestial object whose location has been
 * observed.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#ifndef ObservedStar_h
#define ObservedStar_h

#include "Arduino.h"
#include "Star.h"

class ObservedStar : public Star
{
public:
  ObservedStar(String name, float ra, float dec, float alt, float az, float time);
private:
  float alt;
  float az;
  float time;
};

#endif
