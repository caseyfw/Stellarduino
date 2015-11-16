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

#ifndef StellarduinoUtilities_h
#define StellarduinoUtilities_h

#include <math.h>
#include <avr/pgmspace.h>
#include "Arduino.h"
#include <EEPROM.h>

// Constants.

// EEPROM star catalogue elements.
#define FLOAT_LENGTH 4 // 4 bytes per float number
#define NAME_LENGTH 8 // 8 bytes per star
#define TOTAL_LENGTH 20 // 20 bytes per star total

// Solar day (24h00m00s) / sidereal day (23h56m04.0916s).
const float siderealFraction = 1.002737908;
const float rad2deg = 57.29577951308232;

// Structures.
// TODO: Consider changing the Strings here to char arrays.
struct Star {
  String name;
  float ra;
  float dec;
};
struct ObservedStar {
  String name;
  float ra;
  float dec;
  float alt;
  float az;
  float time;
};
struct CatalogueStar {
  String name;
  float ra;
  float dec;
  float vmag;
};

// Utility functions.
String rad2hm(float rad, boolean highPrecision = false);
String rad2dm(float rad, boolean highPrecision = false);
String padding(String str, int length);

void fillVectorWithT(float* v, float e, float az);
void fillVectorWithC(float* v, ObservedStar star, float initialTime);
void fillVectorWithProduct(float* v, float* a, float* b);

void fillMatrixWithVectors(float* m, float* a, float* b, float* c);
void fillMatrixWithProduct(float* m, float* a, float* b, int aRows, int aCols, int bCols);

void fillStarWithCVector(float* star, float* v, float initialTime);

void copyMatrix(float* recipient, float* donor);
void invertMatrix(float* m);

void loadCatalogueStar(int i, CatalogueStar star);
String readStringFromEEPROM(int offset, int maxLength = NAME_LENGTH);
String readFloatFromEEPROM(int offset);
#endif
