<?php

/**
 * altaztoradec.php
 * PHP test file to play with the Stellarduino PHP class.
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * License: MIT, http://opensource.org/licenses/MIT
 */

require_once('Stellarduino.class.php');

header('Content-Type: text/plain;');

$m22 = new StellarObject(new HMS(18, 36, 24, 'e'), DMS::withDMS(23, 54, 0, 's'), 'M22');
$spica = new StellarObject(new HMS(13, 25, 11.6, 'e'), DMS::withDMS(11, 9, 41, 's'), 'M22');

list($alt, $az) = Stellarduino::getAltAz($m22, -27.43, 153.03);
echo 'Viewed from Brisbane, Australia, M22 is at '.$alt.' alt, '.$az.' az.'."\n";

list($alt, $az) = Stellarduino::getAltAz($spica, -27.43, 153.03);
echo 'Viewed from Brisbane, Australia, Spica is at '.$alt.' alt, '.$az.' az.'."\n";

$dec = Stellarduino::getRADec($alt, $az, -27.43, 153.03);
echo 'Spica has a listed Dec of '.$spica->getDec()->getDecimalDeg().', we calculated '.$dec."\n";