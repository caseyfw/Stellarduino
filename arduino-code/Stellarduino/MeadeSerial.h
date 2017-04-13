/**
 * MeadeSerial.cpp
 *
 * Interface for communicating with a PC over serial, implementing the Meade
 * Autostar protocol.
 *
 * See:
 * http://www.weasner.com/etx/autostar/2010/AutostarSerialProtocol2007oct.pdf
 *
 * Version: 0.4 Better Alignment
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * Website: http://www.caseyfulton.com/stellarduino
 * License: MIT, http://opensource.org/licenses/MIT
 */

#ifndef MeadeSerial_h
#define MeadeSerial_h

#include "Arduino.h"
#include "StellarduinoUtilities.h"

#define WAITING_FOR_START 1
#define WAITING_FOR_END 2
#define START_CHAR ':'
#define END_CHAR '#'
#define GET_RA "GR"
#define GET_DEC "GD"
#define CHANGE_PRECISION "U"

class MeadeSerial
{
public:
  void begin(float obs[2], boolean highPrecision = true, unsigned int baud =
  	9600);
  boolean available();
  void processSerial();
private:
  void _processCommand();
  float * _obs;
  unsigned int _state;
  String _command;
  char _character;
  boolean _highPrecision;
};

#endif
