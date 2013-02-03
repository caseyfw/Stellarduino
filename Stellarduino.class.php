<?php

/**
 * Stellarduino.class.php
 * A set of PHP classes that provide stellar coordinate conversion utility methods.
 * This software was created as a prototype prior to authoring the Arduino port.
 * Author: Casey Fulton, casey AT caseyfulton DOT com
 * License: MIT, http://opensource.org/licenses/MIT
 */


class Stellarduino {
    private static $daysToBeginningOfYearSinceJ2k = array(
        '1998' => -731.5, '1999' => -366.5, '2000' => -1.5,   '2001' => 364.5,
        '2002' => 729.5,  '2003' => 1094.5, '2004' => 1459.5, '2005' => 1825.5,
        '2006' => 2190.5, '2007' => 2555.5, '2008' => 2920.5, '2009' => 3286.5,
        '2010' => 3651.5, '2011' => 4016.5, '2012' => 4381.5, '2013' => 4747.5,
        '2014' => 5112.5, '2015' => 5477.5, '2016' => 5842.5, '2017' => 6208.5,
        '2018' => 6573.5, '2019' => 6938.5, '2020' => 7303.5, '2021' => 7669.5
    );
    
    private static $daysToBeginningOfMonth = array(
      'normal' => array(
        '1' => 0,
        '2' => 31,
        '3' => 59,
        '4' => 90,
        '5' => 120,
        '6' => 151,
        '7' => 181,
        '8' => 212,
        '9' => 243,
        '10' => 273,
        '11' => 304,
        '12' => 334
      ),
      'leap' => array(
        '1' => 0,
        '2' => 31,
        '3' => 60,
        '4' => 91,
        '5' => 121,
        '6' => 152,
        '7' => 182,
        '8' => 213,
        '9' => 244,
        '10' => 274,
        '11' => 305,
        '12' => 335
      )
    );
    
    /*
     * array getAltAz(StellarObject $s, double $lat, double $long)
     * 
     * Calculate the altitude and azimuth of a StellarObject from a given latitude and longitude.
     * Returns an array containing calculated altitude and azimuth values as decimal degrees.
     */
    public static function getAltAz($s, $lat, $long) {
        // get current UTC time
        $now = new DateTime('now', new DateTimezone('UTC'));
        
        // pull out hour portion for readability
        $nowTime = $now->format('G') + $now->format('i') / 60 + $now->format('s') / 3600;
        
        // pull decimal degree versions out of StellarObject s
        $ra = $s->getRA()->getDecimalDeg();
        $dec = $s->getDec()->getDecimalDeg();
        
        // determine time (in days) since j2k
        $timeSinceJ2k =
            self::$daysToBeginningOfYearSinceJ2k[$now->format('Y')] +
            self::$daysToBeginningOfMonth[$now->format('Y') % 4 == 0 ? 'leap' : 'normal'][$now->format('n')] +
            $now->format('j') +
            ($nowTime) / 24;
        
        // determine local sidereal time
        $lst = 100.46 + 0.985647 * $timeSinceJ2k + $long + 15 * $nowTime;
        if ($lst < 0) $lst += 360; // correct lst value to bring into range 0 <= lst < 360
        if ($lst >= 360) $lst -= 360;
        
        // determine hour angle of StellarObject
        $ha = $lst - $ra ;
        
        // determine alt and az
        $alt = rad2deg(asin(self::dsin($dec) * self::dsin($lat) + self::dcos($dec) * self::dcos($lat) * self::dcos($ha)));
        $az = rad2deg(acos((self::dsin($dec) - self::dsin($alt) * self::dsin($lat)) / (self::dcos($alt) * self::dcos($lat))));
        
        if (self::dsin($ha) >= 0) $az = 360 - $az;
        
        // return array(alt,az);
        return array($alt,$az);
    }
    
    private static function dsin($deg) { return sin(deg2rad($deg)); }
    private static function dcos($deg) { return cos(deg2rad($deg)); }
    private static function dasin($deg) { return asin(deg2rad($deg)); }
    private static function dacos($deg) { return acos(deg2rad($deg)); }
    
}

class DMS {
    private $d;
    private $m;
    private $s;
    private $hemisphere; // north/east = 1, south/west = -1
    
    public function __construct($d, $m, $s, $hemisphere) {
        $this->d = abs($d);
        $this->m = $m;
        $this->s = $s;
        $this->hemisphere = ($hemisphere === 'n' || $hemisphere === 'e' ? 1 : -1);
    }
    
    public function getDecimalDeg() {
        return ($this->d + $this->m / 60 + $this->s / 3600) * $this->hemisphere;
    }
}

class HMS {
    private $h;
    private $m;
    private $s;
    private $hemisphere; // north/east = 1, south/west = -1
    
    public function __construct($h, $m, $s, $hemisphere) {
        $this->h = abs($h);
        $this->m = $m;
        $this->s = $s;
        $this->hemisphere = ($hemisphere === 'n' || $hemisphere === 'e' ? 1 : -1);
    }
    
    public function getDecimalHours() {
        return ($this->h + $this->m / 60 + $this->s / 3600) * $this->hemisphere;
    }
    
    public function getDecimalDeg() {
        return ($this->h + $this->m / 60 + $this->s / 3600) * 15 * $this->hemisphere;
    }
}

class StellarObject {
    private $ra;
    private $dec;
    private $name;
    
    public function __construct($ra, $dec, $name) {
        $this->ra = $ra;
        $this->dec = $dec;
        $this->name = $name;
    }
    
    public function getRA() {
        return $this->ra;
    }
    
    public function getDec() {
        return $this->dec;
    }
    
    public function getName() {
        return $this->name;
    }
}