/**
 * StarChecker.ino
 * This sketch displays the contents of an Arduino's EEPROM assuming it has
 * been modified by StarLoader.ino.
 *
 * Unlike StarLoader, this sketch uses only the serial interface to dump the
 * contents of the EEPROM - no LCD or buttons are needed.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include <EEPROM.h>
#include <math.h>

#define FLOAT_LENGTH 4 // 4 bytes per float number
#define NAME_LENGTH 8 // 8 bytes per star
#define TOTAL_LENGTH 20 // 20 bytes per star total
#define NUM_OF_STARS 50 // 20 x 50 = 1000 bytes, ~ the size of the Uno EEPROM.

struct CatalogueStar
{
 char name[NAME_LENGTH + 1]; // Add 1 for the null terminator '/0'.
 float ra;   // Right ascension.
 float dec;  // Declination.
 float vmag; // Apparent magnitude.
};

CatalogueStar star;

char name[NAME_LENGTH + 1];
float* vmag;
float* ra;
float* dec;

/**
 * Adds padding to the beginning or end of a string, using the optionally
 * specified character.
 */
String padding(String str, int length, char character = ' ',
  boolean padOnLeft = true)
{
  while(str.length() < length) {
    if (padOnLeft) {
      str = character + str;
    } else {
      str = str + character;
    }
  }
  return str;
}

/**
 * Converts a float to a string that looks like HH:MM:SS.
 */
String rad2hms(float rad) {
  if (rad < 0) rad = rad + 2.0 * M_PI;
  float hours = rad * 24.0 / (2.0 * M_PI);
  float minutes = (hours - floor(hours)) * 60.0;

  return padding((String)int(floor(hours)), 2, '0') + ":" +
    padding((String)int(floor(minutes)), 2, '0') + ":" +
    padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2, '0');
}

/**
 * Converts a float to a string that looks like +DEG:MM:SS.
 */
String rad2dms(float rad) {
  float degs = abs(rad) * 360.0 / (2.0 * M_PI);
  float minutes = (degs - floor(degs)) * 60.0;
  String sign = "+";
  if (rad < 0) sign = "-";

  return sign + padding((String)int(floor(degs)), 2, '0') + "*" +
    padding((String)int(floor(minutes)), 2, '0') + ":" +
    padding((String)int(floor((minutes - floor(minutes)) * 60.0)), 2, '0');
}

/**
 * Reads a float value from the EEPROM.
 */
float readFloat(int offset) {
  // make a regular four-byte float to hold the value
  float value;
  // make a pointer to byte, and initialise it to point to the first byte of the
  // float
  byte *p = (byte*)(void*)&value;
  for (int i = 0; i < sizeof(value); i++) {
    // assign whatever byte is in EEPROM to the byte p points to
    *p = EEPROM.read(offset + i);
    // move p up to the next byte
    p++;
  }
  return value;
}

void setup()
{
  Serial.begin(9600);

  Serial.println("Ready to read star catalogue from EEPROM. Continue? (y/n)");

  while (!Serial.available());
  if (Serial.read() != 'y') {
    while(true);
  }

  Serial.println("### Beginning dump of EEPROM.");
  Serial.println("No  Name      Magnitude  Right Ascension  Declination");

  for (int i = 0; i < NUM_OF_STARS; i++) {
    for (int c = 0; c < NAME_LENGTH; c++) {
      // fetch name
      star.name[c] = EEPROM.read(i * TOTAL_LENGTH + c);
      if (star.name[c] == (char) 0xFF) {
        star.name[c] = '\0';
        break;
      }
    }

    // fetch floats from next 12 bytes
    star.vmag = readFloat(i * TOTAL_LENGTH + NAME_LENGTH);
    star.ra   = readFloat(i * TOTAL_LENGTH + NAME_LENGTH + FLOAT_LENGTH);
    star.dec  = readFloat(i * TOTAL_LENGTH + NAME_LENGTH + FLOAT_LENGTH * 2);

    // print the star's details with padding so it looks nice
    Serial.print(padding((String) (i + 1), 2) + "  ");
    Serial.print(padding(star.name, 10, ' ', false));
    if (star.vmag > 0) {
      Serial.print(" ");
    }
    Serial.print(star.vmag);
    Serial.print("      ");
    Serial.print(padding(rad2hms(star.ra), 17, ' ', false));
    Serial.println(rad2dms(star.dec));
  }

  Serial.println("### Finished.");
}

void loop()
{
  // do nothing
}
