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
#include <LiquidCrystal.h>
#include "RTClib.h"

// Constants.

// Buttons.
#define OK_BTN A0
#define UP_BTN A1
#define DOWN_BTN A2

// EEPROM star catalogue elements.
#define FLOAT_LENGTH 4 // 4 bytes per float number
#define NAME_LENGTH 8 // 8 bytes per star
#define TOTAL_LENGTH 20 // 20 bytes per star total

#define LAT_ADDR 1000 // The EEPROM address of stored viewing latitude
#define LONG_ADDR 1004 // The EEPROM address of stored viewing longitude

// Solar day (24h00m00s) / sidereal day (23h56m04.0916s).
const float siderealFraction = 1.002737909;
const float rad2deg = 57.295779513;
const float milliRadsPerDay = 542867210.54;

// Structures.
struct Star {
  char name[NAME_LENGTH + 1];
  float ra;
  float dec;
};
struct ObservedStar {
  char name[NAME_LENGTH + 1];
  float ra;
  float dec;
  float alt;
  float az;
  float time; // NOTE: Not sure if this is necessary.
};
struct CatalogueStar {
  char name[NAME_LENGTH + 1];
  float ra;
  float dec;
  float vmag;
};

// Utility functions.
String rad2hms(float rad, boolean highPrecision = false);
String rad2dms(float rad, boolean highPrecision = false);
String padding(String str, int length);
void die();

// LCD interaction functions.
int lcdChoose(LiquidCrystal lcd, char* question, const char answers[][10],
  int answersCount);
void lcdDatePrompt(LiquidCrystal lcd, DateTime d);
void lcdCoordPrompt(LiquidCrystal lcd, char* question, float* value);
void lcdChooseCatalogueStars(LiquidCrystal lcd, ObservedStar* stars);
int waitForButton();
bool inArray(int needle, int* haystack, int count);

// Star catalogue EEPROM functions.
void loadCatalogueStar(int i, CatalogueStar& star);
void loadNameFromEEPROM(int offset, char* name);
void loadFloatFromEEPROM(int offset, float* value);

// Coordinate geometry functions.
float getJulianDate(int year, int month, int day);
float getSiderealTime(float julianDate, float hour = 0.0, float longitude = 0.0);
void celestialToEquatorial(float ra, float dec, float latV, float longV,
  float lst, float* obs);

// Matrix translation functions.
void fillVectorWithT(float* v, float e, float az);
void fillVectorWithC(float* v, ObservedStar star, float initialTime);
void fillVectorWithProduct(float* v, float* a, float* b);

void fillMatrixWithVectors(float* m, float* a, float* b, float* c);
void fillMatrixWithProduct(float* m, float* a, float* b, int aRows, int aCols,
  int bCols);

void fillStarWithCVector(float* star, float* v, float initialTime);

// Generic matrix functions.
void copyMatrix(float* recipient, float* donor);
void invertMatrix(float* m);

#endif
