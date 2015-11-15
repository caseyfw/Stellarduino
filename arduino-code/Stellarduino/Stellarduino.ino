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

// Viewing location expressed as radians.
float viewingCoords[] = {2.4190437966, -0.6096260544};

// Alignment stars.
ObservedStar alignmentStar1;
ObservedStar alignmentStar2;

// Meade serial connection.
MeadeSerial meade;

// Display.
// TODO: Make provision for I2C LCD.
LiquidCrystal lcd(6, 7, 8, 9, 10, 11);

// Real Time Clock.
RTC_DS1307 rtc;

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

// Viewing location
float latV, longV;

// Viewing coords in radians.
float altV, azV;

// Current time UTC.
unsigned long time;

void setup()
{
  // Calculate encoder multipliers based on steps per revolution.
  altMultiplier = 2.0 * M_PI / (float)ALT_SPR;
  azMultiplier = -2.0 * M_PI / (float)AZ_SPR;
  
  lcd.begin(16, 2);
  lcd.clear();
  pinMode(OK_BTN, INPUT);
  pinMode(UP_BTN, INPUT);
  pinMode(DOWN_BTN, INPUT);

  lcd.print("Starting algnmnt");
  Serial.begin(9600);

  if (rtc.begin() && rtc.isrunning()) {
    autoSelectAlignmentStars();
  } else
  {
    manuallySelectAlignmentStars();
  }

  doAlignment();

  calculateTransforms();

  if (DEBUG) {
    Serial.begin(9600);
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
  meade.begin(obs, false, 9600);
}

void loop()
{
  altT = altMultiplier * altEncoder.read();
  azT = azMultiplier * azEncoder.read();

  fillVectorWithT(obsTVector, altT, azT);
  fillMatrixWithProduct(obsCVector, inverseTransformMatrix, obsTVector, 3, 3, 1);
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
    while(Serial.available() == 0)
    {
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
 * Selects alignment stars for the user from the EEPROM star catalogue based on viewing coordinates and time.
 */
void autoSelectAlignmentStars()
{
  lcd.setCursor(0,1);
  lcd.print(rtc.now().unixtime());
  
  // determine utc time (gmt)
  
  // determine viewing lat/long


  // foreach catalogue star
    // calculate alt/az


  while(true) {
      lcd.setCursor(0,1);
      lcd.print(rtc.now().unixtime());
      delay(10);
  }
}

void manuallySelectAlignmentStars()
{
  alignmentStar1 =
  {
    "Arcturus",
    3.73352834160889,
    0.334797783763812,
    -0.04
  };
  alignmentStar2 =
  {
    "Rigel K",
    3.83797175293031,
    -1.06177589858756,
    -0.01
  };
}  

void doAlignment() {
  // Set initial time - actual time not necessary, just the difference!
  initialTime = (float)millis() / 86400000.0f * 2.0 * M_PI;

  // ask user to point scope at first star
  lcd.print("Point: ");
  lcd.print(alignmentStar1.name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // wait for button press
  while(digitalRead(OK_BTN) == LOW);
  alignmentStar1.time = (float)millis() / 86400000.0f * 2.0 * M_PI;
  alignmentStar1.alt = altMultiplier * altEncoder.read();
  alignmentStar1.az = azMultiplier * azEncoder.read();

  lcd.clear();
  lcd.print("Alt set: ");
  lcd.print(alignmentStar1.alt * rad2deg, 3);
  lcd.setCursor(0,1);
  lcd.print("Az set: ");
  lcd.print(alignmentStar1.az * rad2deg, 3);

  delay(2000);

  // ask user to point scope at second star
  lcd.clear();
  lcd.print("Point: ");
  lcd.print(alignmentStar2.name);
  lcd.setCursor(0,1);
  lcd.print("Then press OK");

  // wait for button press
  while(digitalRead(OK_BTN) == LOW);
  alignmentStar2.time = (float)millis() / 86400000.0f * 2.0 * M_PI;
  alignmentStar2.az = azMultiplier * azEncoder.read();
  alignmentStar2.alt = altMultiplier * altEncoder.read();

  lcd.clear();
  lcd.print("Alt set: ");
  lcd.print(alignmentStar2.alt * rad2deg, 3);
  lcd.setCursor(0,1);
  lcd.print("Az set: ");
  lcd.print(alignmentStar2.az * rad2deg, 3);

  delay(2000);
}

void calculateTransforms() {
  // calculate vectors for alignment stars
  fillVectorWithT(firstTVector, alignmentStar1.alt, alignmentStar1.az);
  fillVectorWithT(secondTVector, alignmentStar2.alt, alignmentStar2.az);

  // calculate third's vectors
  fillVectorWithProduct(thirdTVector, firstTVector, secondTVector);

  // calculate celestial vectors for alignment stars
  fillVectorWithC(firstCVector, alignmentStar1, initialTime);
  fillVectorWithC(secondCVector, alignmentStar2, initialTime);

  // calculate third's vector
  fillVectorWithProduct(thirdCVector, firstCVector, secondCVector);

  fillMatrixWithVectors(telescopeMatrix, firstTVector, secondTVector, thirdTVector);
  fillMatrixWithVectors(celestialMatrix, firstCVector, secondCVector, thirdCVector);

  copyMatrix(inverseMatrix, celestialMatrix);
  invertMatrix(inverseMatrix);

  fillMatrixWithProduct(transformMatrix, telescopeMatrix, inverseMatrix, 3, 3, 3);
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
