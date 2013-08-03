<?php

header("Content-type: text/plain");

// ST test

$nowGMT = new DateTime('now', new DateTimezone('GMT'));
$nowUTC = new DateTime('now', new DateTimezone('UCT'));


$daysToBeginningOfYearSinceJ2000 = array(
  '1998' => -731.5,
  '2000' => -1.5,   '2001' => 364.5,  '2002' => 729.5,  '2003' => 1094.5,
  '2004' => 1459.5, '2005' => 1825.5, '2006' => 2190.5, '2007' => 2555.5,
  '2008' => 2920.5, '2009' => 3286.5, '2010' => 3651.5, '2011' => 4016.5,
  '2012' => 4381.5, '2013' => 4747.5, '2014' => 5112.5, '2015' => 5477.5,
  '2016' => 5842.5, '2017' => 6208.5, '2018' => 6573.5, '2019' => 6938.5,
  '2020' => 7303.5, '2021' => 7669.5
);

$daysToBeginningOfMonth = array(
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
    '12' => 334,
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
    '12' => 335,
  )
);

$startApprox = microtime();

$daysSinceJ2k = $daysToBeginningOfYearSinceJ2000[$nowUTC->format('Y')] +
                $daysToBeginningOfMonth[$nowUTC->format('Y') % 4 == 0 ? 'leap' : 'normal'][$nowUTC->format('n')] + 
                $nowUTC->format('j') +
                ($nowUTC->format('G') + ($nowUTC->format('i') / 60) + ($nowUTC->format('s') / 3600)) / 24;
$siderealTime = 100.46 + 0.985647 * $daysSinceJ2k + (15 * ($nowUTC->format('G') + ($nowUTC->format('i') / 60) + ($nowUTC->format('s') / 3600)));

if ($siderealTime < 0) $siderealTime += 360;
if ($siderealTime > 360) $siderealTime = fmod($siderealTime, 360);

$siderealTime = $siderealTime / 360 * 24;

$endApprox = microtime() - $startApprox;

echo "ST using weird approximation equations: ".$siderealTime.", took: $endApprox\n";



$time = $nowUTC;

$startTaki = microtime();

$gmt = $time->setTimezone(new DateTimeZone('GMT'));

$year = (integer)$gmt->format('Y');
$month = (integer)$gmt->format('n');
$day = (integer)$gmt->format('j');
$hour =  $gmt->format('G') + $gmt->format('i') / 60 + $gmt->format('s') / 3600;

if ((integer)$gmt->format('n') < 3) {
    $year = $year - 1;
    $month = $month + 12;
}
$gr = 2 - floor($year / 100) + floor(floor($year / 100) / 4);

$jd = floor(365.25 * $year) + floor(30.6001 * ($month + 1)) + $day + 1720994.5 + $gr;
$jd2 = $jd + $hour / 24;

$time = ($jd - 2415020) / 36525;
$ss = 6.6460656 + 2400.051 * $time + 0.00002581 * ($time ^ 2);
$st = ($ss / 24 - floor($ss / 24)) * 24;

$sa = $st + $hour * 1.002737908;
if ($sa < 0) $sa = $sa + 24;
if ($sa > 24) $sa = $sa - 24;

$endTaki = microtime() - $startTaki;

echo "ST using Taki's equations: ".$sa.", took: $endTaki\n";

