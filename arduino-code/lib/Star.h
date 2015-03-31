/*
  Star.h - Basic star class, root class for all stars, used to store alignment stars.
  Created by Casey Fulton, March 31, 2015.
  Released into the public domain.
*/
#ifndef Star_h
#define Star_h

#include "Arduino.h"

class Star
{
public:
	Star(char[] name, float ra, float dec);
};

#endif
