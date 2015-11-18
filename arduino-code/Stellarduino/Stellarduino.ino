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

#define DEBUG true

// Buttons.
#define OK_BTN A0
#define UP_BTN A1
#define DOWN_BTN A2

// Encoder steps per revolution of scope (typically 4 * CPR * gearing).
#define ALT_SPR 10000
#define AZ_SPR 10000

// The number of stars in the EEPROM catalogue.
#define CATALOGUE_STARS 50

// The number of stars to use during alignment - currently immutable.
#define ALIGNMENT_STARS 2

// Viewing location expressed as radians.
float viewingCoords[] = {2.4190437966, -0.6096260544};

// Alignment stars - loaded from EEPROM.
ObservedStar alignmentStars[2];

// Temp location to put catalogue stars while calculating their suitability.
CatalogueStar catalogueStar;

// Meade serial connection.
MeadeSerial meade;

// Display.
// TODO: Make provision for I2C LCD.
LiquidCrystal lcd(6, 7, 8, 9, 10, 11);

// Real Time Clock.
RTC_DS1307 rtc;

// Real Time Clock date/time object - may come from being manually entered.
DateTime initialDate;

// Initial time as radians.
float initialTime;

// calculation vectors
float firstTVector[3];
float secondTVector[3];
float thirdTVector[3];

float firstCVector[3];
float secondCVector[3];
float thirdCVector[3];

float obsTVector[3];
float obsCVector[3];

// Final observed star coordinates.
float obs[2];

// Matricies.
float telescopeMatrix[9];
float celestialMatrix[9];
float inverseMatrix[9];
float transformMatrix[9];
float inverseTransformMatrix[9];

// Encoders.
Encoder altEncoder(2, 4);
Encoder azEncoder(3, 5);

// Handy modifiers to convert encoder ticks to radians.
float altMultiplier;
float azMultiplier;

// Telescope coords in radians.
float altT, azT;

// Viewing coords in radians.
float altV, azV;

// Sidereal time when the sketch started, expressed in radians.
// TODO: This could replace initialTime, or at least inform it.
float initialSiderealTime;

void setup()
{
  // Calculate encoder multipliers based on steps per revolution.
  altMultiplier = 2.0 * M_PI / (float)ALT_SPR;
  azMultiplier = -2.0 * M_PI / (float)AZ_SPR;

  // Set initial datetime object.
  if (rtc.begin() && rtc.isrunning()) {
    initialDate = rtc.now();
  }

  lcd.begin(16, 2);
  lcd.clear();
  pinMode(OK_BTN, INPUT);
  pinMode(UP_BTN, INPUT);
  pinMode(DOWN_BTN, INPUT);

  if (DEBUG) {
    Serial.begin(9600);
    delay(5000);
  }

  if (rtc.begin() && rtc.isrunning()) {
    Serial.println("Auto selecting stars.");
    autoSelectAlignmentStars();
  } else {
    Serial.println("Manually selecting stars.");
    manuallySelectAlignmentStars();
  }

  Serial.println("Starting alignment.");
  lcd.print("Starting algnmnt");

  doAlignment();

  calculateTransforms();

  if (DEBUG) {
    Serial.println("Telescope matrix:");
    printMatrix(telescopeMatrix);

    Serial.println("Celestial matrix:");
    printMatrix(celestialMatrix);

    Serial.println("Inverse Celestial matrix:");
    printMatrix(inverseMatrix);

    Serial.println("Transform matrix:");
    printMatrix(transformMatrix);

    Serial.println("Inverse Transform matrix:");
    printMatrix(inverseTransformMatrix);
  }

  clearScreen();
//  meade.begin(obs, false, 9600);
}

void loop()
{
  altT = altMultiplier * altEncoder.read();
  azT = azMultiplier * azEncoder.read();

  fillVectorWithT(obsTVector, altT, azT);
  fillMatrixWithProduct(obsCVector, inverseTransformMatrix, obsTVector,
    3, 3, 1);
  fillStarWithCVector(obs, obsCVector, initialTime);

  if (DEBUG) {
    Serial.println("Observed vector:");
    printVector(obsTVector);
    Serial.println("Transformed celestial vector:");
    printVector(obsCVector);
    Serial.println("Celestial coordinates:");
    Serial.print(obs[0]);
    Serial.print(",");
    Serial.println(obs[1]);

    // wait for input from serial before continuing
    while(Serial.available() == 0) {
      // do nothing
    }
    Serial.read();
  }

  lcd.setCursor(5,0);
  lcd.print(rad2hm(obs[0]));
  lcd.print(" ");
  lcd.setCursor(5,1);
  lcd.print(rad2dm(obs[1]));
  lcd.print(" ");

  // if there's a serial request waiting, process it
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
  // float hour = initialDate.hour() + initialDate.minute() / 60.0 +
  //   initialDate.second() / 3600.0;

  // Calculate approximate Julian date.
  // float julian = getJulianDate(initialDate.year(), initialDate.month(), initialDate.day(), hour);

  // OMFG TEST REMOVE ME
  double hour = 7 + 25 / 60.0 + 0 / 3600.0;
  double julian = getJulianDate(2015, 11, 18, hour);

  Serial.print("Julian: ");
  Serial.print(julian, 5);
  Serial.println();

  // Calculate initial local sidereal time.
  initialSiderealTime = getSiderealTime(julian, hour, viewingCoords[1]);

  Serial.print("Initial sidereal time: ");
  Serial.print(initialSiderealTime, 5);
  Serial.println();

  // Alignment star counter.
  int n = 0;
  // foreach catalogue star
  for (int i = 0; i < CATALOGUE_STARS; i++) {
    loadCatalogueStar(i, catalogueStar);

    celestialToEquatorial(
      catalogueStar.ra,
      catalogueStar.dec,
      viewingCoords[0],
      viewingCoords[1],
      initialSiderealTime + (float)millis() / milliRadsPerDay,
      obs
    );

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
  lcd.print("Insuff. alignmnt");
  lcd.setCursor(0, 1);
  lcd.print("stars visible.");
  die();
}

void manuallySelectAlignmentStars()
{
  alignmentStars[0] = (ObservedStar) {
    "Arcturus",
    3.73352834160889,
    0.334797783763812,
    0.0,
    0.0
  };
  alignmentStars[1] = (ObservedStar) {
    "Rigel K",
    3.83797175293031,
    -1.06177589858756,
    0.0,
    0.0
  };
}

void doAlignment() {
  // Set initial time - actual time not necessary, just the difference!
  initialTime = (float)millis() / 86400000.0f * 2.0 * M_PI;

  // ask user to point scope at first star
  lcd.print("Point: ");
  lcd.print(alignmentStars[0].name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // wait for button press
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

  // ask user to point scope at second star
  lcd.clear();
  lcd.print("Point: ");
  lcd.print(alignmentStars[1].name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // wait for button press
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
  // calculate vectors for alignment stars
  fillVectorWithT(firstTVector, alignmentStars[0].alt, alignmentStars[0].az);
  fillVectorWithT(secondTVector, alignmentStars[1].alt, alignmentStars[1].az);

  // calculate third's vectors
  fillVectorWithProduct(thirdTVector, firstTVector, secondTVector);

  // calculate celestial vectors for alignment stars
  fillVectorWithC(firstCVector, alignmentStars[0], initialTime);
  fillVectorWithC(secondCVector, alignmentStars[1], initialTime);

  // calculate third's vector
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
  // apparently I deleted the print matrix function, so I'm adding this back in
  // so debug doesn't die.

  // TODO: rewrite printMatrix.
}

void printVector(float* v)
{
  // apparently I deleted the print vector function, so I'm adding this back in
  // so debug doesn't die.

  // TODO: rewrite printVector.
}
