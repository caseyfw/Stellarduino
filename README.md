Stellarduino
============

Stellarduino is an Arduino-powered telescope computer, offering two star alignment, Push-To navigation and Meade Autostar compatible serial output for displaying telescope orientation on a PC in [Stellarium](http://www.stellarium.org).

For information on how to get started, check out www.caseyfulton.com/stellarduino/.

Arduino Sketches
================

Stellarduino - The main one. Takes input from rotary encoders and after an alignment procedure, does matrix coordinate translation to respond to Meade Autostar serial requests over USB (presumably coming from Stellarium).

StarLoader - One-time use sketch to load a catalogue of the 50 brightest stars into your Arduino's persistant EEPROM memory via the USB serial connection. These are then used to perform the alignment process.

StarChecker - Used to check if the star catalogue has been loaded correctly.