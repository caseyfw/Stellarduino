/**
 * Stellarduino.ino
 * The base Arduino sketch that makes up the heart of Stellarduino.
 *
 * This software is pretty dodgy, but accomplishes PushTo so long as you
 * preselect alignment stars below.
 *
 * Software Requirements
 * TLB's Encoder library: http://www.pjrc.com/teensy/td_libs_Encoder.html
 * Adafruit's RTC library: https://github.com/adafruit/RTClib
 *
 * Hardware Requirements
 * To run this program, you'll need an Arduino Uno or better, a 16x2 LCD,
 * display, a push button, and a 220k ohm resistor.
 *
 * For more information, including a setup guide, head to
 * www.caseyfulton.com/stellarduino
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include <EEPROM.h>
#include <Encoder.h>
#include <LiquidCrystal.h>
#include <math.h>
#include <Wire.h>
#include "RTClib.h"
#include "StellarduinoUtilities.h"
#include "MeadeSerial.h"

#define DEBUG false

// Encoder steps per revolution of scope (typically 4 * CPR * gearing).
#define ALT_SPR 10000
#define AZ_SPR 10000

// Meade serial connection.
MeadeSerial meade;

// Display.
// TODO: Make provision for I2C LCD.
LiquidCrystal lcd(6, 7, 8, 9, 10, 11);

// Encoders.
Encoder altEncoder(2, 4);
Encoder azEncoder(3, 5);

// Real Time Clock.
RTC_DS1307 rtc;

// Real Time Clock date/time object - may come from being manually entered.
DateTime initialDate;

// Initial time as radians.
float initialTime;

// Sidereal time when the sketch started, expressed in radians.
// TODO: This could replace initialTime, or at least inform it.
float initialSiderealTime;

// Alignment stars - loaded from EEPROM.
ObservedStar alignmentStars[2];

// Temp location to put catalogue stars while calculating their suitability.
CatalogueStar catalogueStar;

// Handy modifiers to convert encoder ticks to radians.
float altMultiplier;
float azMultiplier;

// Unprocessed telescope orientation in radians from encoders.
float altT, azT;

// Viewing coordinates in radians.
float latitude, longitude;

// Calculation vectors.
float firstTVector[3];
float secondTVector[3];
float thirdTVector[3];

float firstCVector[3];
float secondCVector[3];
float thirdCVector[3];

float obsTVector[3];
float obsCVector[3];

// Matricies.
float telescopeMatrix[9];
float celestialMatrix[9];
float inverseMatrix[9];
float transformMatrix[9];
float inverseTransformMatrix[9];

// Final observed star coordinates [ra, dec] in radians.
float obs[2];

void setup()
{
  // Calculate encoder multipliers based on steps per revolution.
  altMultiplier = 2.0 * M_PI / (float)ALT_SPR;
  azMultiplier = -2.0 * M_PI / (float)AZ_SPR;

  // Set initial datetime object.
  if (rtc.begin() && rtc.isrunning()) {
    initialDate = rtc.now();
  }

  // Attempt to fetch viewing location from EEPROM.
  loadFloatFromEEPROM(LAT_ADDR, &latitude);
  loadFloatFromEEPROM(LONG_ADDR, &longitude);

  lcd.begin(16, 2);
  lcd.clear();
  pinMode(OK_BTN, INPUT);
  pinMode(UP_BTN, INPUT);
  pinMode(DOWN_BTN, INPUT);

  if (DEBUG) {
    Serial.begin(9600);
    delay(5000);
  }

  const char starSelectionOptions[][10] = {"Auto", "Semi-auto", "Manually"};

  switch (lcdChoose(lcd, "Star selection?", starSelectionOptions, 3)) {
    // Automatic alignment star selection.
    case 0:
      // Check if RTC is working and viewing coordinates make sense.
      if (rtc.isrunning() &&
        latitude >= M_PI * -0.5 && latitude <= M_PI * 0.5 &&
        longitude >= M_PI * -1.0 && longitude <= M_PI ) {
        autoSelectAlignmentStars();
      } else {
        lcd.clear();
        lcd.print("Error: date or");
        lcd.setCursor(0,1);
        lcd.print("coords not set.");
        die();
      }
      break;

    // Semi-automatic alignment star selection.
    case 1:
      // Determine date and time.
      lcdDatePrompt(lcd, initialDate);
      lcdCoordPrompt(lcd, "Enter latitude", &latitude);
      lcdCoordPrompt(lcd, "Enter longitude", &longitude);

      // Once data is collected, star selection can be performed.
      autoSelectAlignmentStars();
      break;

    // Manual alignment star selection.
    case 2:
      lcdChooseCatalogueStars(lcd, alignmentStars);
      break;
  }

  lcd.clear();
  lcd.print("Star 1: ");
  lcd.print(alignmentStars[0].name);
  lcd.setCursor(0,1);
  lcd.print("Star 2: ");
  lcd.print(alignmentStars[1].name);
  delay(5000);

  doAlignment();

  calculateTransforms();

  clearScreen();
  meade.begin(obs, false, 9600);
}

void loop()
{
  // Read encoder values.
  altT = altMultiplier * altEncoder.read();
  azT = azMultiplier * azEncoder.read();

  // Use transformation matrix to convert to RA/Dec.
  fillVectorWithT(obsTVector, altT, azT);
  fillMatrixWithProduct(obsCVector, inverseTransformMatrix, obsTVector,
    3, 3, 1);
  fillStarWithCVector(obs, obsCVector, initialTime);

  // Refresh LCD.
  lcd.setCursor(5,0);
  lcd.print(rad2hms(obs[0]));
  lcd.print(" ");
  lcd.setCursor(5,1);
  lcd.print(rad2dms(obs[1]));
  lcd.print(" ");

  // If there's a serial request waiting, process it.
  if (Serial.available()) {
    meade.processSerial();
  }
}

/**
 * Selects alignment stars for the user from the EEPROM star catalogue based on
 * viewing coordinates and time.
 */
void autoSelectAlignmentStars()
{
  // Hours, minutes and seconds in decimal since program started running.
  float hour = initialDate.hour() + initialDate.minute() / 60.0 +
    initialDate.second() / 3600.0;

  // Calculate approximate current Julian day.
  float julianDate = getJulianDate(initialDate.year(), initialDate.month(),
    initialDate.day());

  // Calculate initial local sidereal time.
  initialSiderealTime = getSiderealTime(julianDate, hour, longitude);

  // Alignment star counter.
  uint8_t n = 0;
  for (uint8_t i = 0; i < CATALOGUE_STARS; i++) {
    loadCatalogueStar(i, catalogueStar);

    celestialToEquatorial(
      catalogueStar.ra,
      catalogueStar.dec,
      latitude,
      longitude,
      initialSiderealTime + (float)millis() / 13713441.095,
      obs
    );
    // TODO: Figure out the milliRadsPerSiderealDay issue.

    // If catalogue star is higher than 25 degrees above the horizon.
    if (obs[0] > 0.436332313) {
      // Copy catalogue star to alignment star.
      strcpy(alignmentStars[n].name, catalogueStar.name);
      alignmentStars[n].ra = catalogueStar.ra;
      alignmentStars[n].dec = catalogueStar.dec;
      alignmentStars[n].alt = obs[0];
      alignmentStars[n].az = obs[1];
      n++;

      // If both alignment stars have been selected, return.
      if (n >= ALIGNMENT_STARS) {
        return;
      }
    }
  }

  // If we get to here, insufficient alignment stars have been selected. Error!
  lcd.clear();
  lcd.print("Insuff. alignmnt");
  lcd.setCursor(0, 1);
  lcd.print("stars visible.");
  die();
}

void doAlignment() {
  // Set initial time - actual time not necessary, just the difference!
  initialTime = (float)millis() / 86400000.0f * 2.0 * M_PI;

  // Ask user to point scope at first star.
  lcd.clear();
  lcd.print("Point: ");
  lcd.print(alignmentStars[0].name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // Wait for button press.
  while(digitalRead(OK_BTN) == LOW);
  alignmentStars[0].time = (float)millis() / 86400000.0f * 2.0 * M_PI;
  alignmentStars[0].alt = altMultiplier * altEncoder.read();
  alignmentStars[0].az = azMultiplier * azEncoder.read();

  lcd.clear();
  lcd.print("Alt set: ");
  lcd.print(alignmentStars[0].alt * rad2deg, 3);
  lcd.setCursor(0,1);
  lcd.print("Az set: ");
  lcd.print(alignmentStars[0].az * rad2deg, 3);

  delay(2000);

  // Ask user to point scope at second star.
  lcd.clear();
  lcd.print("Point: ");
  lcd.print(alignmentStars[1].name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // Wait for button press.
  while(digitalRead(OK_BTN) == LOW);
  alignmentStars[1].time = (float)millis() / 86400000.0f * 2.0 * M_PI;
  alignmentStars[1].az = azMultiplier * azEncoder.read();
  alignmentStars[1].alt = altMultiplier * altEncoder.read();

  lcd.clear();
  lcd.print("Alt set: ");
  lcd.print(alignmentStars[1].alt * rad2deg, 3);
  lcd.setCursor(0,1);
  lcd.print("Az set: ");
  lcd.print(alignmentStars[1].az * rad2deg, 3);

  delay(2000);
}

void calculateTransforms() {
  // Calculate vectors for alignment stars.
  fillVectorWithT(firstTVector, alignmentStars[0].alt, alignmentStars[0].az);
  fillVectorWithT(secondTVector, alignmentStars[1].alt, alignmentStars[1].az);

  // Calculate third's vectors.
  fillVectorWithProduct(thirdTVector, firstTVector, secondTVector);

  // Calculate celestial vectors for alignment stars.
  fillVectorWithC(firstCVector, alignmentStars[0], initialTime);
  fillVectorWithC(secondCVector, alignmentStars[1], initialTime);

  // Calculate third's vector.
  fillVectorWithProduct(thirdCVector, firstCVector, secondCVector);

  fillMatrixWithVectors(telescopeMatrix, firstTVector, secondTVector,
    thirdTVector);
  fillMatrixWithVectors(celestialMatrix, firstCVector, secondCVector,
    thirdCVector);

  copyMatrix(inverseMatrix, celestialMatrix);
  invertMatrix(inverseMatrix);

  fillMatrixWithProduct(transformMatrix, telescopeMatrix, inverseMatrix,
    3, 3, 3);
  copyMatrix(inverseTransformMatrix, transformMatrix);
  invertMatrix(inverseTransformMatrix);
}

void clearScreen()
{
  lcd.clear();
  lcd.print("RA: ");
  lcd.setCursor(0, 1);
  lcd.print("Dec:  ");
}

void printMatrix(float* m)
{
  // Apparently I deleted the print matrix function, so I'm adding this back in
  // so debug doesn't die.

  // TODO: rewrite printMatrix.
}

void printVector(float* v)
{
  // Apparently I deleted the print vector function, so I'm adding this back in
  // so debug doesn't die.

  // TODO: rewrite printVector.
}
