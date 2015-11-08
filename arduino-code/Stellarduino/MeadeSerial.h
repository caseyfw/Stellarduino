/*
  MeadeSerial.h - Interface for interacting with Meade telescopes over serial.
  Created by Casey Fulton, March 31, 2015.
  Released into the public domain.
*/

#ifndef MeadeSerial_h
#define MeadeSerial_h

#include "Arduino.h"
#include "StellarduinoUtils.h"

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
  MeadeSerial(float obs[2], boolean highPrecision = true, unsigned int baud = 9600);
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
