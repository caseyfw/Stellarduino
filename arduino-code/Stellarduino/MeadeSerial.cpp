/*
  MeadeSerial.cpp - Interface for interacting with Meade telescopes over serial.
  Created by Casey Fulton, November 7, 2015.
  Released into the public domain.
*/

#include "MeadeSerial.h"

MeadeSerial::MeadeSerial(float obs[2], boolean highPrecision, unsigned int baud)
{
  _obs = obs;
  _highPrecision = highPrecision;
  _state = WAITING_FOR_START;
  Serial.begin(baud);
}

boolean MeadeSerial::available()
{
  return Serial.available();
}

void MeadeSerial::processSerial()
{
  _character = Serial.read();
  if (_state == WAITING_FOR_START)
  {
    if (_character == START_CHAR)
    {
      _state = WAITING_FOR_END;
      _command = "";
//      lcd.setCursor(15,0);
//      lcd.print((char)126);
    }
  } else if (_state == WAITING_FOR_END)
  {
    if (_character == END_CHAR)
    {
//      lcd.setCursor(15,0);
//      lcd.print((char)127);
      _processCommand();
      _state = WAITING_FOR_START;
//      lcd.setCursor(15,0);
//      lcd.print(" ");
    } else
    {
      _command += _character;
    }
  }
}

void MeadeSerial::_processCommand()
{
  if (_command == GET_RA)
  {
    Serial.print("#" + StellarduinoUtils::rad2hm(_obs[0], _highPrecision) + "#");
//    lcd.setCursor(15,1);
//    lcd.print('R');
  } else if (_command == GET_DEC)
  {
    Serial.print("#" + StellarduinoUtils::rad2dm(_obs[1], _highPrecision) + "#");
//    lcd.setCursor(15,1);
//    lcd.print('D');
  } else if (_command == CHANGE_PRECISION)
  {
    _highPrecision = !_highPrecision;
//    clearScreen();
//    lcd.setCursor(15,1);
//    lcd.print('P');
  }
}

