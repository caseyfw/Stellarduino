/**
 * StellarduinoUtils.h
 * Some type defines and utility functions used by Stellarduino.
 */

#include <math.h>
#include <avr/pgmspace.h>
#include "Arduino.h"

#ifndef AlignmentStarStruct
#define AlignmentStarStruct

// Structure for storing alignment star data.
struct AlignmentStar {
  String name;
  float ra;
  float dec;
  float vmag;
};

#endif

#ifndef SimpleStarStruct
#define SimpleStarStruct

// Structure for storing star data.
struct SimpleStar {
  float time;
  float ra;
  float dec;
  float alt;
  float az;
  String name;
  float vmag;
};

#endif

#ifndef StellarduinoUtils_h
#define StellarduinoUtils_h

class StellarduinoUtils
{
public:
  static String rad2hm(float rad, boolean highPrecision = false);
  static String rad2dm(float rad, boolean highPrecision = false);
  static String padding(String str, int length);
  boolean available();
  void processSerial();
private:
  void _processCommand();
  float * _obs;
  unsigned int _state;
  String _command;
  char _character;
  boolean _highPrecision;
};

#endif
