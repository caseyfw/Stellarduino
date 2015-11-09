/**
 * StarLoader.ino
 * This sketch uploads a catalogue of the 50 brightest stars to an Arduino's
 * EEPROM.
 *
 * WARNING: EEPROM is not like regular flash memory, it has a limited life span,
 * and will actually "wear out" after ~100,000 erase/write cycles. There is no
 * advertised limit on the number of times the EEPROM can be read, but
 * nonetheless some caution is advised when running this sketch.
 *
 * Each star is represented using 20 bytes. The first 8 bytes is used for the
 * star's name. The next 4 is a float representing apparent magnitude, the last
 * 8 are two floats, representing right ascension and declination in decimal
 * radians. The stars are uploaded in order of vmag, with the first being the
 * brightest (Sirius, -1.46) and the last being the dimmest (Debeb K, 2.04).
 *
 * This star catalogue is designed to be consumed by Stellarduino, in order to
 * provide automatic alignment star selection, based on their visibility and
 * apparent magnitude.
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#include <EEPROM.h>

#define FLOAT_LENGTH 4 // 4 bytes per float number.
#define NAME_LENGTH 8 // 8 bytes per star.
#define TOTAL_LENGTH 20 // 20 bytes per star total.
#define NUM_OF_STARS 50 // 20 x 50 = 1000 bytes, ~ the size of the Uno EEPROM.

struct CatalogueStar
{
  char name[NAME_LENGTH + 1]; // Add 1 for the null terminator '/0'.
  float vmag; // Apparent magnitude.
  float ra;   // Right ascension.
  float dec;  // Declination.
};

CatalogueStar stars[] = {
  { "Sirius",   -1.46, 1.76779309390854,  -0.291751177018097 },
  { "Canopus",  -0.72, 1.67530518796327,  -0.919715793748845 },
  { "Arcturus", -0.04, 3.73352834160889,   0.334797783763812 },
  { "Rigel K",  -0.01, 3.83797175293031,  -1.06177589858756 },
  { "Vega",      0.03, 4.87356286460115,   0.676901709701945 },
  { "Capella",   0.08, 1.38182080203521,   0.802817518959714 },
  { "Rigel",     0.12, 1.37243238510052,  -0.143146087484402 },
  { "Procyon",   0.38, 2.00408158580771,   0.0911934534167037 },
  { "Achernar",  0.46, 0.426362119646565, -0.998968286199821 },
  { "Betelgse",  0.50, 1.54972874828228,   0.129275568067858 },
  { "Hadar",     0.61, 3.68187386795507,  -1.0537085989339 },
  { "Altair",    0.77, 5.19577246113495,   0.15478161583103 },
  { "Aldbaran",  0.85, 1.20392811802569,   0.288139315093831 },
  { "Antares",   0.96, 4.31710099362884,  -0.461324458259779 },
  { "Spica",     0.98, 3.51331869544372,  -0.194802985206623 },
  { "Pollux",    1.14, 2.03031970222935,   0.489147915418655 },
  { "Fomalht",   1.16, 6.01113938223019,  -0.517005309535209 },
  { "Mimosa",    1.25, 3.34981043335272,  -1.04176278983136 },
  { "Deneb",     1.25, 5.41676750546352,   0.790289933439843 },
  { "Acrux",     1.33, 3.2576497766422,   -1.10128821359799 },
  { "Regulus",   1.35, 2.65452216479469,   0.20886743009561 },
  { "Adhara",    1.50, 1.82659614529032,  -0.505660669397246 },
  { "Gacrux",    1.63, 3.2775756189358,   -0.996815713455695 },
  { "Shaula",    1.63, 4.59723361077915,  -0.647585026405252 },
  { "Bellatrx",  1.64, 1.41865452145751,   0.110823559364829 },
  { "El Nath",   1.65, 1.42371597628829,   0.499295065764278 },
  { "Miaplcds",  1.68, 2.41379035550816,  -1.21679507312234 },
  { "Alnilan",   1.70, 1.46700741394297,  -0.0209778879816096 },
  { "Al Na'ir",  1.74, 5.7955112253515,   -0.819626009283782 },
  { "Alioth",    1.77, 3.37733573009771,   0.976681401279216 },
  { "VelaGam2",  1.78, 2.13599211623239,  -0.826180690252382 },
  { "Mirfak",    1.79, 0.891528726329137,  0.870240557591617 },
  { "Dubhe",     1.79, 2.89606118886027,   1.07775535751693 },
  { "Wezen",     1.84, 1.86921126785984,  -0.460650567243037 },
  { "Kaus Aus",  1.85, 4.81785777264166,  -0.600126615161439 },
  { "Avior",     1.86, 2.19262805045961,  -1.03864058972501 },
  { "Alkaid",    1.86, 3.61082442298847,   0.860680031800137 },
  { "Sargas",    1.87, 4.61342881179661,  -0.750452793263073 },
  { "Menkalmn",  1.90, 1.56873829271859,   0.784481865540151 },
  { "Atria",     1.92, 4.40113132490715,  -1.2047619975572 },
  { "Alhena",    1.93, 1.73534451423188,   0.286219452916637 },
  { "Peacock",   1.94, 5.34789972206191,  -0.990212551118983 },
  { "VelaDelt",  1.96, 2.28945019071399,  -0.954840544945231 },
  { "Mirzam",    1.98, 1.66984376184557,  -0.313388411606015 },
  { "Alphard",   1.98, 1.98356669489156,   0.556556409640125 },
  { "Castor",    1.98, 2.47656403093822,  -0.151121272538653 },
  { "Hamal",     2.00, 0.55489834685073,   0.40949787574917 },
  { "Nunki",     2.02, 0.662403356568364,  1.55795361238231 },
  { "Polaris",   2.02, 4.95352803316336,  -0.458963415632776 },
  { "Deneb K",   2.04, 0.190182710825649,  0.924791792990062 }
};

/**
 * Writes the 20-byte details of a star to EEPROM at given memory offset.
 */
boolean writeStar(int offset, char name[NAME_LENGTH + 1], float vmag,
  float ra, float dec)
{
  // For each character in the name, until end of string or max length.
  for (int i = 0; name[i] != '\0' && i < NAME_LENGTH; i++)
  {
    EEPROM.write(offset + i, (byte) name[i]);
  }

  // Split floats into four bytes and write to EEPROM.
  writeFloat(offset + NAME_LENGTH, vmag);
  writeFloat(offset + NAME_LENGTH + FLOAT_LENGTH, ra);
  writeFloat(offset + NAME_LENGTH + FLOAT_LENGTH * 2, dec);

}

/**
 * Writes the details of a star to EEPROM.
 */
boolean writeFloat(int offset, float number)
{
  // cast float pointer to byte pointer
  byte* b = (byte*) &number;
  for (int i = 0; i < 4; i++)
  {
    // dereference pointer to get byte value
    EEPROM.write(offset + i, *b);

    // increment the byte pointer
    b++;
  }
}

void setup()
{
  Serial.begin(9600);

  Serial.println("Ready to upload star catalogue to EEPROM. Continue? (y/n)");

  while (!Serial.available());
  if (Serial.read() != 'y')
  {
    while(true);
  }

  Serial.println("Starting upload.");

  for (int i = 0; i < NUM_OF_STARS; i++)
  {
    Serial.print((String) (i + 1) + "/" + (String) NUM_OF_STARS + " ");
    Serial.print(stars[i].name);
    Serial.println();

    writeStar(i * TOTAL_LENGTH, stars[i].name, stars[i].vmag, stars[i].ra, stars[i].dec);

    delay(200);
  }

  Serial.print("Upload succeeded!");
}

void loop()
{
  // do nothing
}

