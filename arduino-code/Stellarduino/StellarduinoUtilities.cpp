/**
 * StellarduinoUtilities.cpp
 *
 * Some type defines and utility functions used by Stellarduino.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include "StellarduinoUtilities.h"

String rad2hms(float rad, boolean highPrecision)
{
  if (rad < 0) rad = rad + 2.0 * M_PI;
  float hours = rad * 12.0 / M_PI;
  float minutes = (hours - floor(hours)) * 60.0;

  if (highPrecision) {
    return padding((String)int(floor(hours)), (uint8_t)2) + ":" +
      padding((String)int(floor(minutes)), (uint8_t)2) + ":" +
      padding((String)int(floor((minutes - floor(minutes)) * 60.0)), (uint8_t)2);
  } else {
    return padding((String)int(floor(hours)), (uint8_t)2) + ":" +
      padding((String)int(floor(minutes)), (uint8_t)2) + "." +
      (String)int(floor((minutes - floor(minutes)) * 10.0));
  }
}

String rad2dms(float rad, boolean highPrecision)
{
  float degs = abs(rad) * 360.0 / (2.0 * M_PI);
  float minutes = (degs - floor(degs)) * 60.0;
  String sign = "+";
  if (rad < 0) sign = "-";

  if (highPrecision) {
    return sign + padding((String)int(floor(degs)), (uint8_t)2) + "*" +
      padding((String)int(floor(minutes)), (uint8_t)2) + "'" +
      padding((String)int(floor((minutes - floor(minutes)) * 60.0)), (uint8_t)2);
  } else {
    return sign + padding((String)int(floor(degs)), (uint8_t)2) + "*" +
      padding((String)int(floor(minutes)), (uint8_t)2);
  }
}

String padding(String str, uint8_t length)
{
  while(str.length() < length) {
    str = "0" + str;
  }
  return str;
}

/**
 * Returns true if needle is present in array haystack. Only for integers, and
 * O(N) complexity (i.e. shitty). Why is this functionality not in the standard
 * Arduino library?
 */
bool inArray(uint8_t needle, uint8_t* haystack, uint8_t count)
{
  for (uint8_t i = 0; i < count; i++) {
    if (haystack[i] == needle) {
      return true;
    }
  }
  return false;
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

uint8_t lcdChoose(LiquidCrystal lcd, char* question, const char answers[][10],
  uint8_t answersCount)
{
  uint8_t selection = 0;
  uint8_t button;

  lcd.clear();
  lcd.print(question);

  while (true) {
    lcd.setCursor(0, 1);
    lcd.print("* ");
    lcd.print(answers[selection]);
    lcd.print("              ");

    button = waitForButton();

    if (button == OK_BTN) return selection;
    if (button == UP_BTN) selection--;
    if (button == DOWN_BTN) selection++;

    // Prevent selection from wrapping the answers array.
    if (selection < 0) {
      selection = selection + answersCount;
    } else if (selection >= answersCount) {
      selection = selection % answersCount;
    }
  }
}

void lcdDatePrompt(LiquidCrystal lcd, DateTime d)
{
  char question[] = "Enter UTC Date";
  char answer[] = "YYYY-MM-DD HH:MM";
  uint8_t skipPositions[] = {4, 7, 10, 13};
  char characters[] = {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9'};

  lcdPrompt(lcd, question, answer, (uint8_t)16, skipPositions, (uint8_t)4,
    characters, (uint8_t)10);

  // Build DateTime object from answer entered into LCD.
  d = DateTime(
    atoi(&answer[0]),
    atoi(&answer[5]),
    atoi(&answer[8]),
    atoi(&answer[11]),
    atoi(&answer[14]),
    0
  );
}

void lcdCoordPrompt(LiquidCrystal lcd, char* question, float* value)
{
  char answer[] = "SNNN.NNNNNNN";
  uint8_t skipPositions[] = {4};
  char characters[] = {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '+',
    '-'};

  lcdPrompt(lcd, question, answer, (uint8_t)12, skipPositions, (uint8_t)1,
    characters, (uint8_t)12);

  // Convert answer text to float.
  *value = atof(answer) * deg2rad;
}

/**
 * Prompts for an answer which the user enters one character at a time using
 * only the LCD and three buttons.
 */
void lcdPrompt(LiquidCrystal lcd, char* question, char* answer, uint8_t
  answerLength, uint8_t* skipPositions, uint8_t skipsCount, char* characters,
  uint8_t charactersCount)
{
  uint8_t button;
  uint8_t cursorPosition = 0;
  int8_t currentCharacter = 0;

  // Print the question to the display.
  lcd.clear();
  lcd.print(question);

  // Print the answer to the display as placeholder text.
  lcd.setCursor(0, 1);
  lcd.print(answer);
  lcd.setCursor(0, 1);

  // Enable the cursor.
  lcd.cursor();

  while (true) {
    // Write current character to screen, then reset cursor on top of it.
    lcd.print(characters[currentCharacter]);
    lcd.setCursor(cursorPosition, 1);

    button = waitForButton();

    if (button == OK_BTN) {
      // Add selected character to answer output string.
      answer[cursorPosition] = characters[currentCharacter];

      // Move cursor along, skipping cells if necessary.
      cursorPosition++;
      while (inArray(cursorPosition, skipPositions, skipsCount)) {
        cursorPosition++;
      }
      lcd.setCursor(cursorPosition, 1);

      // Reset currentCharacter.
      // TODO: Remember char when returning to a position that's already set.
      currentCharacter = 0;

      // If at end of answer, break out of loop.
      if (cursorPosition >= answerLength) {
        break;
      }
    } else if (button == UP_BTN) {
      currentCharacter--;

    } else if (button == DOWN_BTN) {
      currentCharacter++;
    }

    // Prevent currentCharacter from wrapping the characters array.
    if (currentCharacter < 0) {
      currentCharacter = currentCharacter + charactersCount;
    } else if (currentCharacter >= charactersCount) {
      currentCharacter = currentCharacter % charactersCount;
    }
  }

  lcd.noCursor();
}

void lcdChooseCatalogueStars(LiquidCrystal lcd, ObservedStar* stars)
{
  CatalogueStar catalogueStar;
  uint8_t button;
  uint8_t starIndex = 0;
  int8_t currentStar = 0;

  while (true) {
    // Load star from catalogue.
    loadCatalogueStar(currentStar, catalogueStar);

    // Print the question to the display.
    lcd.setCursor(0, 0);
    lcd.print("Select ");
    lcd.print(ALIGNMENT_STARS - starIndex);
    lcd.print(" stars ");
    lcd.setCursor(0, 1);
    lcd.print(catalogueStar.name);
    lcd.print("        ");

    button = waitForButton();

    if (button == OK_BTN) {
      // Copy selected star's details across.
      strcpy(stars[starIndex].name, catalogueStar.name);
      stars[starIndex].ra = catalogueStar.ra;
      stars[starIndex].dec = catalogueStar.dec;

      // Move index to next star.
      starIndex++;

      // If at end of answer, break out of loop.
      if (starIndex >= ALIGNMENT_STARS) {
        break;
      }
    } else if (button == UP_BTN) {
      currentStar--;

    } else if (button == DOWN_BTN) {
      currentStar++;
    }

    // Prevent currentStar from wrapping the catalogue.
    if (currentStar < 0) {
      currentStar = currentStar + ALIGNMENT_STARS;
    } else if (currentStar >= ALIGNMENT_STARS) {
      currentStar = currentStar % ALIGNMENT_STARS;
    }
  }
}

/**
 * Waits for, then returns the pin number of the button that was pressed.
 *
 * TODO: Refactor this, it's horrible. There has to be a better way.
 */
 uint8_t waitForButton()
{
  uint8_t button;

  while (true) {
    // Poor man's "wait for button to be pressed".
    while (digitalRead(OK_BTN) == 0 && digitalRead(UP_BTN) == 0 &&
      digitalRead(DOWN_BTN) == 0) {}

    // Poor man's "which button was pressed?".
    button = digitalRead(OK_BTN) ? OK_BTN :
      digitalRead(UP_BTN) ? UP_BTN :
      digitalRead(DOWN_BTN) ? DOWN_BTN : -1;

    // Poor man's debounce.
    delay(400);
    return button;
  }
}

void loadCatalogueStar(uint8_t i, CatalogueStar& star)
{
  uint8_t offset = i * TOTAL_LENGTH;
  loadNameFromEEPROM(offset, star.name);
  loadFloatFromEEPROM(offset + NAME_LENGTH, &star.ra);
  loadFloatFromEEPROM(offset + NAME_LENGTH + FLOAT_LENGTH, &star.dec);
  loadFloatFromEEPROM(offset + NAME_LENGTH + FLOAT_LENGTH + FLOAT_LENGTH,
    &star.vmag);
}

/**
 * Reads a name (char array) from the EEPROM into the referenced char array.
 */
void loadNameFromEEPROM(uint8_t offset, char* name)
{
  for (uint8_t c = 0; c < NAME_LENGTH; c++) {
    name[c] = EEPROM.read(offset + c);
    // If the character is blank, replace it with the null terminator.
    if (name[c] == (char) 0xFF) {
      name[c] = '\0';
      break;
    }
  }
}

/**
 * Reads a float value from the EEPROM into the referenced float.
 */
void loadFloatFromEEPROM(uint8_t offset, float* value)
{
  // Make pointer to byte, and make it to point to the first byte of the float.
  byte *p = (byte*)value;

  for (uint8_t i = 0; i < FLOAT_LENGTH; i++) {
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
float getJulianDate(uint16_t year, uint8_t month, uint8_t day)
{
  uint16_t gregorian;

  // Massage year/month to work with approximation formula below.
  if (month < 3) {
      year = year - 1;
      month = month + 12;
  }

  // Approximate the difference between Gregorian and Julian dates.
  gregorian = 2 - floor(year / 100.0) + floor(year / 400.0);

  // Julian date approximation.
  return floor(365.25 * year) + floor(30.6001 * (month + 1)) + day + gregorian +
    1720994.5;
}

float getSiderealTime(float julian, float hour, float longitude)
{
  float s;

  // Julian centuries since J2000.0.
  s = (julian - 2451545.0) / 36525.0;

  // Sidereal time approximation quadratic.
  s = 6.697374558 + 2400.051336 * s + 0.000025862 * s * s;

  // Mod back to 0 < s < 24. This keeps just the fraction part of s / 24.
  s = fmod(s, 24.0);

  // Add hours at the sidereal rate.
  s = s + hour * siderealFraction;

  // Mod back again, just in case it goes over 24. Doing this twice increases
  // accuracy, because we're restricted to single precision floats.
  s = fmod(s, 24.0);

  // Add in viewer's longitude offset, which must be converted from radians.
  s = s + longitude * (12.0 / M_PI);

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
    ha += 2.0 * M_PI;
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

void fillMatrixWithProduct(float* m, float* a, float* b, uint8_t aRows, uint8_t aCols,
  uint8_t bCols)
{
  for (uint8_t i = 0; i < aRows; i++) {
    for (uint8_t j = 0; j < bCols; j++) {
      m[bCols * i + j] = 0;
      for (uint8_t k = 0; k < aCols; k++) {
        m[bCols * i + j] = m[bCols * i + j] + a[aCols * i + k] * b[bCols * k + j];
      }
    }
  }
}

void copyMatrix(float* recipient, float* donor)
{
  for (uint8_t i = 0; i < 9; i++) {
    recipient[i] = donor[i];
  }
}

void invertMatrix(float* m)
{
  float temp;
  uint8_t pivrow;
  uint8_t pivrows[9];
  uint8_t i,j,k;

  for (k = 0; k < 3; k++) {
    temp = 0;
    for (i = k; i < 3; i++) {
      if (abs(m[i * 3 + k]) >= temp) {
        temp = abs(m[i * 3 + k]);
        pivrow = i;
      }
    }
    if (pivrow != k) {
      for (j = 0; j < 3; j++) {
        temp = m[k * 3 + j];
        m[k * 3 + j] = m[pivrow * 3 + j];
        m[pivrow * 3 + j] = temp;
      }
    }

    // Record pivot row swap.
    pivrows[k] = pivrow;

    temp = 1.0 / m[k * 3 + k];
    m[k * 3 + k] = 1.0;

    // Row reduction.
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
