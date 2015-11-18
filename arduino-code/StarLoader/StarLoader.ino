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
 * Star Catalogue Schema
 *
 * Each star is represented using 20 bytes. The first 8 bytes are used for the
 * star's name. The next 8 are two single-precision floats, representing right
 * ascension and declination in decimal radians. The last 4 bytes are another
 * float representing apparent magnitude. The stars are uploaded in order of
 * vmag, with the first being the brightest (Sirius, -1.46) and the last being
 * the dimmest (Debeb K, 2.04).
 *
 * Name                     RA           Dec          Magnitude
 * S  i  r  i  u  s  \0 ?   1.767793     -0.291751    -1.460000
 * 53 69 72 69 75 73 00 ??  3f e2 47 0b  be 95 60 69  bf ba e1 48
 * A  r  c  t  u  r  u  s   3.733528     0.334798     -0.040000
 * 41 72 63 74 75 72 75 73  40 6e f2 21  3e ab 6a 9d  bd 23 d7 0a
 *
 * Note how Sirius, being only 6 characters long is terminated by a null
 * terminator byte, whereas Arcturus at 8 characters long doesn't have one.
 * This allows for maximum space usage, but can result in weird side effects if
 * not accounted for when reading from the catalogue.
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
  float ra;   // Right ascension.
  float dec;  // Declination.
  float vmag; // Apparent magnitude.
};

CatalogueStar stars[] = {
  { "Sirius",    1.76779309390854,  -0.291751177018097, -1.46 },
  { "Canopus",   1.67530518796327,  -0.919715793748845, -0.72 },
  { "Arcturus",  3.73352834160889,   0.334797783763812, -0.04 },
  { "Rigel K",   3.83797175293031,  -1.06177589858756,  -0.01 },
  { "Vega",      4.87356286460115,   0.676901709701945,  0.03 },
  { "Capella",   1.38182080203521,   0.802817518959714,  0.08 },
  { "Rigel",     1.37243238510052,  -0.143146087484402,  0.12 },
  { "Procyon",   2.00408158580771,   0.0911934534167037, 0.38 },
  { "Achernar",  0.426362119646565, -0.998968286199821,  0.46 },
  { "Betelgse",  1.54972874828228,   0.129275568067858,  0.50 },
  { "Hadar",     3.68187386795507,  -1.0537085989339,    0.61 },
  { "Altair",    5.19577246113495,   0.15478161583103,   0.77 },
  { "Aldbaran",  1.20392811802569,   0.288139315093831,  0.85 },
  { "Antares",   4.31710099362884,  -0.461324458259779,  0.96 },
  { "Spica",     3.51331869544372,  -0.194802985206623,  0.98 },
  { "Pollux",    2.03031970222935,   0.489147915418655,  1.14 },
  { "Fomalht",   6.01113938223019,  -0.517005309535209,  1.16 },
  { "Mimosa",    3.34981043335272,  -1.04176278983136,   1.25 },
  { "Deneb",     5.41676750546352,   0.790289933439843,  1.25 },
  { "Acrux",     3.2576497766422,   -1.10128821359799,   1.33 },
  { "Regulus",   2.65452216479469,   0.20886743009561,   1.35 },
  { "Adhara",    1.82659614529032,  -0.505660669397246,  1.50 },
  { "Gacrux",    3.2775756189358,   -0.996815713455695,  1.63 },
  { "Shaula",    4.59723361077915,  -0.647585026405252,  1.63 },
  { "Bellatrx",  1.41865452145751,   0.110823559364829,  1.64 },
  { "El Nath",   1.42371597628829,   0.499295065764278,  1.65 },
  { "Miaplcds",  2.41379035550816,  -1.21679507312234,   1.68 },
  { "Alnilan",   1.46700741394297,  -0.0209778879816096, 1.70 },
  { "Al Na'ir",  5.7955112253515,   -0.819626009283782,  1.74 },
  { "Alioth",    3.37733573009771,   0.976681401279216,  1.77 },
  { "VelaGam2",  2.13599211623239,  -0.826180690252382,  1.78 },
  { "Mirfak",    0.891528726329137 , 0.870240557591617,  1.79 },
  { "Dubhe",     2.89606118886027,   1.07775535751693,   1.79 },
  { "Wezen",     1.86921126785984,  -0.460650567243037,  1.84 },
  { "Kaus Aus",  4.81785777264166,  -0.600126615161439,  1.85 },
  { "Avior",     2.19262805045961,  -1.03864058972501,   1.86 },
  { "Alkaid",    3.61082442298847,   0.860680031800137,  1.86 },
  { "Sargas",    4.61342881179661,  -0.750452793263073,  1.87 },
  { "Menkalmn",  1.56873829271859,   0.784481865540151,  1.90 },
  { "Atria",     4.40113132490715,  -1.2047619975572,    1.92 },
  { "Alhena",    1.73534451423188,   0.286219452916637,  1.93 },
  { "Peacock",   5.34789972206191,  -0.990212551118983,  1.94 },
  { "VelaDelt",  2.28945019071399,  -0.954840544945231,  1.96 },
  { "Mirzam",    1.66984376184557,  -0.313388411606015,  1.98 },
  { "Alphard",   1.98356669489156,   0.556556409640125,  1.98 },
  { "Castor",    2.47656403093822,  -0.151121272538653,  1.98 },
  { "Hamal",     0.55489834685073,   0.40949787574917,   2.00 },
  { "Nunki",     0.662403356568364 , 1.55795361238231,   2.02 },
  { "Polaris",   4.95352803316336,  -0.458963415632776,  2.02 },
  { "Deneb K",   0.190182710825649 , 0.924791792990062,  2.04 }
};

/**
 * Writes the 20-byte details of a star to EEPROM at given memory offset.
 */
boolean writeStar(int offset, char name[NAME_LENGTH + 1], float ra, float dec,
  float vmag)
{
  // For each character in the name, until max length.
  for (int i = 0; i < NAME_LENGTH; i++) {
    EEPROM.write(offset + i, (byte) name[i]);

    // If current character is the null terminator, break out.
    if (name[i] == '\0') {
      break;
    }
  }

  // Split floats into four bytes and write to EEPROM.
  writeFloat(offset + NAME_LENGTH, ra);
  writeFloat(offset + NAME_LENGTH + FLOAT_LENGTH, dec);
  writeFloat(offset + NAME_LENGTH + FLOAT_LENGTH * 2, vmag);

}

/**
 * Writes the details of a star to EEPROM.
 */
boolean writeFloat(int offset, float number)
{
  // cast float pointer to byte pointer
  byte* b = (byte*) &number;
  for (int i = 0; i < 4; i++) {
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
  if (Serial.read() != 'y') {
    while(true);
  }

  Serial.println("Starting upload.");

  for (int i = 0; i < NUM_OF_STARS; i++) {
    Serial.print((String) (i + 1) + "/" + (String) NUM_OF_STARS + " ");
    Serial.print(stars[i].name);
    Serial.println();

    writeStar(i * TOTAL_LENGTH, stars[i].name, stars[i].ra,
      stars[i].dec, stars[i].vmag);

    delay(200);
  }

  Serial.print("Upload succeeded!");
}

void loop()
{
  // do nothing
}
