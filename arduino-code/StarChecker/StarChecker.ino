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
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include <EEPROM.h>
#include "EEPROMAnything.h"

#define FLOAT_LENGTH 4 // 4 bytes per float number
#define NAME_LENGTH 9 // 8 bytes per star, 1 for the null terminator
#define TOTAL_LENGTH 20 // 20 bytes per star total
#define NUM_OF_STARS 50 // 20 x 50 = 1000 bytes, roughly the size of the Arduino Uno EEPROM

struct CatalogueStar
{
  char name[NAME_LENGTH - 1];
  float vmag;
  float ra;
  float dec;
};

CatalogueStar star = { 'Arcturus', -0.04, 3.73352834160889,   0.334797783763812 };


void setup()
{
  Serial.begin(9600);

  Serial.print("Length of catalogueStar in bytes: ");
  Serial.println(sizeof(star));

  Serial.println("### Beginning dump of EEPROM.");
  Serial.println("No  Name      Right Ascension  Declination  Magnitude");
  
  /*
  for (int i = 0; i < NUM_OF_STARS; i++)
  {
    writeStar(i * TOTAL_LENGTH, stars[i].name, stars[i].brightness, stars[i].ra, stars[i].dec);
    EEPROM.read(i * TOTAL_LENGTH, 8);
  }
  */

  Serial.println("### Finished.");
}

void loop()
{
  // do nothing
}

/*
boolean readStar(int startIndex)
{
  // for each character in the name, until null termination char or max length is reached...
  for (int i = 0; name[i] != '\0' && i < NAME_LENGTH - 1; i++)
  {
    Serial.println((String) (startIndex + i) + ": " + (String) name[i]);
    EEPROM.read(startIndex + i, (byte) name[i]);
  }
  
  // split floats into four bytes and write to EEPROM
  writeFloat(startIndex + NAME_LENGTH - 1, brightness);
  writeFloat(startIndex + NAME_LENGTH - 1 + FLOAT_LENGTH, ra);
  writeFloat(startIndex + NAME_LENGTH - 1 + FLOAT_LENGTH * 2, dec);
  
}

boolean readFloat(int startIndex)
{
  // cast float pointer to byte pointer
  byte* b = (byte*) &number;
  for (int i = 0; i < 4; i++)
  {
    Serial.print((String) (startIndex + i) + ": ");
    
    // dereference pointer to get byte value
    Serial.println(*b, HEX);
    EEPROM.read(startIndex + i, *b);
    
    // increment the byte pointer
    b++;
  }
}
*/
