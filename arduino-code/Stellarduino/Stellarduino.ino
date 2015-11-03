/**
 * Stellarduino.ino
 * The base Arduino sketch that makes up the heart of Stellarduino
 *
 * This software is pretty dodgy, but accomplishes PushTo so long as you
 * preselect alignment stars below.
 *
 * Version: 0.3 Meade Autostar
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * License: MIT, http://opensource.org/licenses/MIT
 *
 * Choosing new alignment stars
 * Feel free to replace the alignment stars below with ones that are visible
 * from your location/season.
 */

#include <Encoder.h>
#include <LiquidCrystal.h>
#include <math.h>
//#include <Wire.h>
//#include "RTClib.h"
 #include "AlignmentStars.h"

#define WAITING_FOR_START 1
#define WAITING_FOR_END 2
#define START_CHAR ':'
#define END_CHAR '#'
#define GET_RA "GR"
#define GET_DEC "GD"
#define CHANGE_PRECISION "U"

#define DEBUG false

// Structure for storing star data.
typedef struct {
  float time;
  float ra;
  float dec;
  float alt;
  float az;
  String name;
  float vmag;
} Star;

//RTC_DS1307 RTC;

// some function prototypes
void fillVectorWithT(float* v, float e, float az);
void fillVectorWithC(float* v, Star star, float initialTime);
void fillStarWithCVector(float* star, float* v);

// solar day (24h00m00s) / sidereal day (23h56m04.0916s)
const float siderealFraction = 1.002737908;

// initial time as radians
float initialTime;

// alignment stars
AlignmentStar alignmentStar1 = ;
Star alignmentStar2 = {0.0, 3.73352834160889,  0.334797783763812, 0.0, 0.0, "Arcturus", -0.04};

// calculation vectors
float firstTVector[3];
float secondTVector[3];
float thirdTVector[3];

float firstCVector[3];
float secondCVector[3];
float thirdCVector[3];

float obsTVector[3];
float obsCVector[3];

// final observed star coordinates
float obs[2];

// matricies
float telescopeMatrix[9];
float celestialMatrix[9];
float inverseMatrix[9];
float transformMatrix[9];
float inverseTransformMatrix[9];

// encoders
Encoder altEncoder(2, 4);
Encoder azEncoder(3, 5);

// state and Autostar protocol command buffers
int state;
String command;
char character;
boolean highPrecision;

// display
LiquidCrystal lcd(6, 7, 8, 9, 10, 11);

// encoder steps per revolution of scope (typically 4 * CPR * gearing)
const int altSPR = 10000;
const int azSPR = 10000;

// handy modifiers to convert encoder ticks to radians
float altMultiplier, azMultiplier, altT, azT;

const float rad2deg = 57.29577951308232;

// buttons
const int OK_BTN = A0;

void setup()
{
  Serial.begin(9600);
  lcd.begin(16, 2);
  lcd.clear();
  pinMode(OK_BTN, INPUT);

  // setup encoders
  altMultiplier = 2.0 * M_PI / ((float)altSPR);
  azMultiplier = -2.0 * M_PI / ((float)azSPR);

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

  state = WAITING_FOR_START;
  highPrecision = true;
}

void loop()
{
  altT = altMultiplier * altEncoder.read();
  azT = azMultiplier * azEncoder.read();

  fillVectorWithT(obsTVector, altT, azT);
  fillMatrixWithProduct(obsCVector, inverseTransformMatrix, obsTVector, 3, 3, 1);
  fillStarWithCVector(obs, obsCVector);

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
    processSerial();
  }
}

void fillVectorWithT(float* v, float e, float az) {
  v[0] = cos(e) * cos(az);
  v[1] = cos(e) * sin(az);
  v[2] = sin(e);
}

void fillVectorWithC(float* v, Star star, float initialTime) {
  v[0] = cos(star.dec) * cos(star.ra - siderealFraction * (star.time - initialTime));
  v[1] = cos(star.dec) * sin(star.ra - siderealFraction * (star.time - initialTime));
  v[2] = sin(star.dec);
}

void fillStarWithCVector(float* star, float* v)
{
  star[0] = atan(v[1] / v[0]) + siderealFraction * ((float)millis() / 86400000.0f * 2.0 * M_PI - initialTime);
  if(v[0] < 0) star[0] = star[0] + M_PI;
  star[1] = asin(v[2]);
}

void fillVectorWithProduct(float* v, float* a, float* b) {
  float multiplier = 1 / sqrt(
    pow(a[1] * b[2] - a[2] * b[1], 2) +
    pow(a[2] * b[0] - a[0] * b[2], 2) +
    pow(a[0] * b[1] - a[1] * b[0], 2)
  );
  v[0] = multiplier * (a[1] * b[2] - a[2] * b[1]);
  v[1] = multiplier * (a[2] * b[0] - a[0] * b[2]);
  v[2] = multiplier * (a[0] * b[1] - a[1] * b[0]);
}

void fillMatrixWithVectors(float* m, float* a, float* b, float* c)
{
  m[0] = a[0];
  m[1] = b[0];
  m[2] = c[0];
  m[3] = a[1];
  m[4] = b[1];
  m[5] = c[1];
  m[6] = a[2];
  m[7] = b[2];
  m[8] = c[2];
}

void invertMatrix(float* m) {
  float temp;
  int pivrow;
  int pivrows[9];
  int i,j,k;

  for(k = 0; k < 3; k++) {
    temp = 0;
    for(i = k; i < 3; i++) {
      if(abs(m[i * 3 + k]) >= temp) {
        temp = abs(m[i * 3 + k]);
        pivrow = i;
      }
    }
    // should do something here... if(m[pivrow * 3 + k] == 0.0) "singular matrix"
    if(pivrow != k) {
      for(j = 0; j < 3; j++) {
        temp = m[k * 3 + j];
        m[k * 3 + j] = m[pivrow * 3 + j];
        m[pivrow * 3 + j] = temp;
      }
    }

    //record pivot row swap
    pivrows[k] = pivrow;

    temp = 1.0 / m[k * 3 + k];
    m[k * 3 + k] = 1.0;

    // row reduction
    for(j = 0; j < 3; j++) {
      m[k * 3 + j] = m[k * 3 + j] * temp;
    }

    for(i = 0; i < 3; i++) {
      if(i != k) {
        temp = m[i* 3 + k];
        m[i * 3 + k] = 0.0;
        for(j = 0; j < 3; j++) {
          m[i * 3 + j] = m[i * 3 + j] - m[k * 3 + j] * temp;
        }
      }
    }
  }

  for(k = 2; k >= 0; k--) {
    if(pivrows[k] != k) {
      for(i = 0; i < 3; i++) {
        temp = m[i * 3 + k];
        m[i * 3 + k] = m[i * 3 + pivrows[k]];
        m[i * 3 + pivrows[k]] = temp;
      }
    }
  }
}

void fillMatrixWithProduct(float* m, float* a, float* b, int aRows, int aCols, int bCols)
{
  for(int i = 0; i < aRows; i++) {
    for(int j = 0; j < bCols; j++) {
      m[bCols * i + j] = 0;
      for(int k = 0; k < aCols; k++) {
        m[bCols * i + j] = m[bCols * i + j] + a[aCols * i + k] * b[bCols * k + j];
      }
    }
  }
}

void copyMatrix(float* recipient, float* donor)
{
  for(int i = 0; i < 9; i++) {
    recipient[i] = donor[i];
  }
}

void doAlignment() {
  // set initial time - actual time not necessary, just the difference!
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

void processSerial()
{
  character = Serial.read();
  if (state == WAITING_FOR_START)
  {
    if (character == START_CHAR)
    {
      state = WAITING_FOR_END;
      command = "";
      lcd.setCursor(15,0);
      lcd.print((char)126);
    }
  } else if (state == WAITING_FOR_END)
  {
    if (character == END_CHAR)
    {
      lcd.setCursor(15,0);
      lcd.print((char)127);
      processCommand();
      state = WAITING_FOR_START;
      lcd.setCursor(15,0);
      lcd.print(" ");
    } else
    {
      command += character;
    }
  }
}

void processCommand()
{
  if (command == GET_RA)
  {
    Serial.print("#" + rad2hm(obs[0]) + "#");
    lcd.setCursor(15,1);
    lcd.print('R');
  } else if (command == GET_DEC)
  {
    Serial.print("#" + rad2dm(obs[1]) + "#");
    lcd.setCursor(15,1);
    lcd.print('D');
  } else if (command == CHANGE_PRECISION)
  {
    highPrecision = !highPrecision;
    clearScreen();
    lcd.setCursor(15,1);
    lcd.print('P');
  }
}

String rad2hm(float rad) {
  if (rad < 0) rad = rad + 2.0 * M_PI;
  float hours = rad * 24.0 / (2.0 * M_PI);
  float minutes = (hours - floor(hours)) * 60.0;

  if (highPrecision)
  {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + ":" + padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  } else {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + "." + (String)int(floor((minutes - floor(minutes)) * 10.0));
  }
}

String rad2dm(float rad) {
  float degs = abs(rad) * 360.0 / (2.0 * M_PI);
  float minutes = (degs - floor(degs)) * 60.0;
  String sign = "+";
  if (rad < 0) sign = "-";

  if (highPrecision)
  {
    return sign + padding((String)int(floor(degs)), 2) + "*" + padding((String)int(floor(minutes)), 2) + ":" + padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  }
  {
    return sign + padding((String)int(floor(degs)), 2) + "*" + padding((String)int(floor(minutes)), 2);
  }
}

String padding(String str, int length) {
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
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
