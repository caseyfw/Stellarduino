/*
  ObservedStar.h - Stored alignment star class.
  Created by Casey Fulton, March 31, 2015.
  Released into the public domain.
*/
#ifndef ObservedStar_h
#define ObservedStar_h

#include "Arduino.h"
#include "Star.h"

class ObservedStar : public Star
{
public:
	ObservedStar(char[] name, float ra, float dec, float alt, float az, float time);
};

#endif
