/*
  MeadeSerial.h - Interface for interacting with Meade telescopes over serial.
  Created by Casey Fulton, March 31, 2015.
  Released into the public domain.
*/
#ifndef MeadeSerial_h
#define MeadeSerial_h

#include "Arduino.h"

class MeadeSerial
{
public:
	MeadeSerial(unsigned int baud = 9600);
	boolean available();
	void processSerial();
private:
	unsigned int baud;
	void processCommand(String command);
};

#endif
