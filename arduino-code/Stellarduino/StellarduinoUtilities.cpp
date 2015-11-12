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

  if (highPrecision)
  {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + ":" + padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2);
  } else {
    return padding((String)int(floor(hours)), 2) + ":" + padding((String)int(floor(minutes)), 2) + "." + (String)int(floor((minutes - floor(minutes)) * 10.0));
  }
}

String rad2dm(float rad, boolean highPrecision)
{
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

String padding(String str, int length)
{
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
}

void fillVectorWithT(float* v, float e, float az)
{
  v[0] = cos(e) * cos(az);
  v[1] = cos(e) * sin(az);
  v[2] = sin(e);
}

void fillVectorWithC(float* v, ObservedStar star, float initialTime)
{
  v[0] = cos(star.dec) * cos(star.ra - siderealFraction * (star.time - initialTime));
  v[1] = cos(star.dec) * sin(star.ra - siderealFraction * (star.time - initialTime));
  v[2] = sin(star.dec);
}

void fillStarWithCVector(float* star, float* v, float initialTime)
{
  star[0] = atan(v[1] / v[0]) + siderealFraction * ((float)millis() / 86400000.0f * 2.0 * M_PI - initialTime);
  if(v[0] < 0) star[0] = star[0] + M_PI;
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

