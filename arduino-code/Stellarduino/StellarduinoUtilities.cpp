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

#include "StellarduinoUtilities.h"

String rad2hm(float rad, boolean highPrecision)
{
  if (rad < 0) rad = rad + 2.0 * M_PI;
  float hours = rad * 24.0 / (2.0 * M_PI);
  float minutes = (hours - floor(hours)) * 60.0;

  if (highPrecision) {
    return padding((String)int(floor(hours)), 2) + ":" +
      padding((String)int(floor(minutes)), 2) + ":" +
      padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  } else {
    return padding((String)int(floor(hours)), 2) + ":" +
      padding((String)int(floor(minutes)), 2) + "." +
      (String)int(floor((minutes - floor(minutes)) * 10.0));
  }
}

String rad2dm(float rad, boolean highPrecision)
{
  float degs = abs(rad) * 360.0 / (2.0 * M_PI);
  float minutes = (degs - floor(degs)) * 60.0;
  String sign = "+";
  if (rad < 0) sign = "-";

  if (highPrecision) {
    return sign + padding((String)int(floor(degs)), 2) + "*" +
      padding((String)int(floor(minutes)), 2) + ":" +
      padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  } else {
    return sign + padding((String)int(floor(degs)), 2) + "*" +
      padding((String)int(floor(minutes)), 2);
  }
}

String padding(String str, int length)
{
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
}

/**
 * Where bad Arduino programs go to die. Literally does nothing forever.
 */
void die()
{
  while (true) {
    // Do nothing. The end. They're all dead, Jim.
  }
}

void loadCatalogueStar(int i, CatalogueStar star)
{
  int offset = i * TOTAL_LENGTH;
  loadNameFromEEPROM(offset, star.name);
  loadFloatFromEEPROM(offset + NAME_LENGTH, &star.ra);
  loadFloatFromEEPROM(offset + NAME_LENGTH + FLOAT_LENGTH, &star.ra);
  loadFloatFromEEPROM(offset + NAME_LENGTH + FLOAT_LENGTH + FLOAT_LENGTH,
    &star.ra);
}

/**
 * Reads a name (char array) from the EEPROM.
 */
float loadNameFromEEPROM(int offset, char* name)
{
  for (int c = 0; c < NAME_LENGTH; c++) {
    name[c] = EEPROM.read(offset + c);
    // If the character is blank, replace it with the null terminator.
    if (name[c] == (char) 0xFF) {
      name[c] = '\0';
      break;
    }
  }
}

/**
 * Reads a float value from the EEPROM.
 */
float loadFloatFromEEPROM(int offset, float* value)
{
  // Make pointer to byte, and make it to point to the first byte of the float.
  byte *p = (byte*)(void*)&value;

  for (int i = 0; i < FLOAT_LENGTH; i++) {
    // Assign whatever byte is in EEPROM to the byte p points to.
    *p = EEPROM.read(offset + i);
    // Move p up to the next byte.
    p++;
  }
}

/**
 * Approximates the Julian date for the current one. Not valid for dates before
 * 1582 AD.
 */
float getJulianDate(int year, int month, int day, float hour)
{
  float gregorian;

  // Massage year/month to work with Gregorian approximation formula below.
  if (month < 3) {
      year = year - 1;
      month = month + 12;
  }

  // Approximate the difference between Gregorian and Julian dates.
  gregorian = 2 - floor(year / 100.0) + floor(floor(year / 100.0) / 4.0);

  // Julian date approximation.
  return floor(365.25 * year) + floor(30.6001 * month) + day + hour / 24.0 +
    1720994.5 + gregorian;
}

float getSiderealTime(float julian, float hour, float longitude)
{
  float s;

  // Approximation of Julian centuries since 1900.
  s = (julian - 2415020) / 36525.0;

  s = 6.6460656 + 2400.051 * s + 0.00002581 * s * s;
  // This is basically MOD 24.
  s = (s / 24.0 - floor(s / 24.0)) * 24;
  s = s + hour * 1.002737908;

  // Add in viewer's longitude offset (in radians).
  s = s + longitude / (M_PI / 12.0);

  // Massage to make result 0 < sidereal < 24.
  if (s < 0) s = s + 24;
  if (s > 24) s = s - 24;

  // Return in radians.
  return s / 12.0 * M_PI;
}

void celestialToEquatorial(float ra, float dec, float latV, float longV,
  float lst, float* obs)
{
  float ha = lst - ra;
  if (ha < 0) {
    ha += 2 * M_PI;
  }
  obs[0] = asin(sin(dec) * sin(latV) + cos(dec) * cos(latV) * cos(ha));
  obs[1] = acos((sin(dec) - sin(obs[0]) * sin(latV)) / (cos(obs[0]) * cos(latV)));
}

void fillVectorWithT(float* v, float e, float az)
{
  v[0] = cos(e) * cos(az);
  v[1] = cos(e) * sin(az);
  v[2] = sin(e);
}

void fillVectorWithC(float* v, ObservedStar star, float initialTime)
{
  v[0] = cos(star.dec) * cos(star.ra - siderealFraction * (star.time -
    initialTime));
  v[1] = cos(star.dec) * sin(star.ra - siderealFraction * (star.time -
    initialTime));
  v[2] = sin(star.dec);
}

void fillStarWithCVector(float* star, float* v, float initialTime)
{
  star[0] = atan(v[1] / v[0]) + siderealFraction * ((float)millis() /
    milliRadsPerDay - initialTime);
  if (v[0] < 0) star[0] = star[0] + M_PI;
  star[1] = asin(v[2]);
}

void fillVectorWithProduct(float* v, float* a, float* b)
{
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

void fillMatrixWithProduct(float* m, float* a, float* b, int aRows, int aCols,
  int bCols)
{
  for (int i = 0; i < aRows; i++) {
    for (int j = 0; j < bCols; j++) {
      m[bCols * i + j] = 0;
      for (int k = 0; k < aCols; k++) {
        m[bCols * i + j] = m[bCols * i + j] + a[aCols * i + k] * b[bCols * k + j];
      }
    }
  }
}

void copyMatrix(float* recipient, float* donor)
{
  for (int i = 0; i < 9; i++) {
    recipient[i] = donor[i];
  }
}

void invertMatrix(float* m)
{
  float temp;
  int pivrow;
  int pivrows[9];
  int i,j,k;

  for (k = 0; k < 3; k++) {
    temp = 0;
    for (i = k; i < 3; i++) {
      if (abs(m[i * 3 + k]) >= temp) {
        temp = abs(m[i * 3 + k]);
        pivrow = i;
      }
    }
    // should do something here... if (m[pivrow * 3 + k] == 0.0) "singular matrix"
    if (pivrow != k) {
      for (j = 0; j < 3; j++) {
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
    for (j = 0; j < 3; j++) {
      m[k * 3 + j] = m[k * 3 + j] * temp;
    }

    for (i = 0; i < 3; i++) {
      if (i != k) {
        temp = m[i* 3 + k];
        m[i * 3 + k] = 0.0;
        for (j = 0; j < 3; j++) {
          m[i * 3 + j] = m[i * 3 + j] - m[k * 3 + j] * temp;
        }
      }
    }
  }

  for (k = 2; k >= 0; k--) {
    if (pivrows[k] != k) {
      for (i = 0; i < 3; i++) {
        temp = m[i * 3 + k];
        m[i * 3 + k] = m[i * 3 + pivrows[k]];
        m[i * 3 + pivrows[k]] = temp;
      }
    }
  }
}
