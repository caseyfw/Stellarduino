<?php

header("Content-type: text/plain");

$k = 1.002737908; // solar day (24h00m00s) / sidereal day (23h56m04.0916s)

const USE_TAKI_VALUES = false;
const USE_OTHER_TAKI_VALUES = true;

if (USE_TAKI_VALUES) {

    $initialTime = 5.497787;

    $first = array(
        'name' => 'aAnd',
        'time' => 5.619669,
        'ra' => 0.034470,
        'dec' => 0.506809,
        'az' => 1.732239,
        'alt' => 1.463808
    );

    $second = array(
        'name' => 'aUmi',
        'time' => 5.659376,
        'ra' => 0.618501,
        'dec' => 1.557218,
        'az' => 5.427625,
        'alt' => 0.611563
    );

    $test = array(
        'name' => 'bCet',
        'time' => 5.725553,
        'ra' => 0.188132,
        'dec' => -0.314822
    );
    $test = $test;
} elseif (USE_OTHER_TAKI_VALUES) {
    $aand = array(
        'name' => 'A And',
        'time' => 5.619669303,
        'ra' => 0.034470253,
        'dec' => 0.506808708,
        'az' => 1.732177875,
        'alt' => 1.489254544
    );
    $aumi = array(
        'name' => 'A Umi',
        'time' => 5.659375544,
        'ra' => 0.618501054,
        'dec' => 1.557217665,
        'az' => 5.427621054,
        'alt' => 0.636992817
    );
    $acyg = array(
        'name' => 'A Cyg',
        'time' => 5.581053894,
        'ra' => 5.415320337,
        'dec' => 0.789709127,
        'az' => 0.169081123,
        'alt' => 0.873450024
    );
    $alyr = array(
        'name' => 'A Lyr',
        'time' => 5.566145873,
        'ra' => 4.872159329,
        'dec' => 0.676751417,
        'az' => 0.182229743,
        'alt' => 0.469371396
    );
    $epeg = array(
        'name' => 'E Peg',
        'time' => 5.754786876,
        'ra' => 5.688537087,
        'dec' => 0.171600772,
        'az' => 1.037047188,
        'alt' => 0.730507558
    );
    $dcap = array(
        'name' => 'D Cap',
        'time' => 5.589416930,
        'ra' => 5.700754391,
        'dec' => -0.282219740,
        'az' => 1.574359513,
        'alt' => 0.520544449
    );
    $apsa = array(
        'name' => 'A PsA',
        'time' => 5.652466949,
        'ra' => 6.008877726,
        'dec' => -0.517891549,
        'az' => 1.916293476,
        'alt' => 0.414672777
    );
    $bcet = array(
        'name' => 'B Cet',
        'time' => 5.725552611,
        'ra' => 0.188131949,
        'dec' => -0.314822490,
        'az' => 2.276998708,
        'alt' => 0.682877523
    );
    $acet = array(
        'name' => 'A Cet',
        'time' => 5.839726233,
        'ra' => 0.793179423,
        'dec' => 0.070738195,
        'az' => 3.094256151,
        'alt' => 0.905197563
    );
    $aari = array(
        'name' => 'A Ari',
        'time' => 5.734497424,
        'ra' => 0.552542152,
        'dec' => 0.408721204,
        'az' => 3.374183806,
        'alt' => 1.242220642
    );
    $atau = array(
        'name' => 'A Tau',
        'time' => 5.702136110,
        'ra' => 1.201513746,
        'dec' => 0.287804794,
        'az' => 3.806289712,
        'alt' => 0.635230035
    );
    $aaur = array(
        'name' => 'A Aur',
        'time' => 5.693918518,
        'ra' => 1.378737387,
        'dec' => 0.802642016,
        'az' => 4.475105697,
        'alt' => 0.681114741
    );
    $aper = array(
        'name' => 'A Per',
        'time' => 5.718716738,
        'ra' => 0.888518033,
        'dec' => 0.869662660,
        'az' => 4.565295101,
        'alt' => 1.023426167
    );
    $bcas = array(
        'name' => 'B Cas',
        'time' => 5.744751233,
        'ra' => 0.037815467,
        'dec' => 1.031437228,
        'az' => 5.647428260,
        'alt' => 1.146943118
    );
    $first = $bcet;
    $second = $apsa;
    $test = $bcas;
    $initialTime = $first['time'];
} else {
    $rigel = array(
        'name' => 'Rigel Kentaurus',
        'time' => 3.5342917352885173,
        'ra' => 3.8380153861616138,
        'dec' => -1.0615664590773222,
        'az' => 3.6286219332219996,
        'alt' => 0.8136773454165676
    );

    $arcturus = array(
        'name' => 'Arcturus',
        'time' => 3.547381704678475,
        'ra' => 3.7335065249932367,
        'dec' => 0.3347376668673547,
        'az' => 5.402215822825014,
        'alt' => 0.420503146310356
    );
    
    $vega = array(
        'name' => 'Vega',
        'time' => 3.560471674068432,
        'ra' => 4.87356286460115,
        'dec' => 0.6769133452302918,
        'az' => 0.2388240674513685,
        'alt' => 0.3844378565726177
    );
    $fomalhaut = array(
        'name' => 'Fomalhaut',
        'time' => 3.5735616434583894,
        'ra' => 6.011146654435404,
        'dec' => -0.5170116121130637,
        'az' => 1.981399577736996,
        'alt' => 0.357763407837971
    );
    $hadar = array(
        'name' => 'Hadar',
        'time' => 3.586651612848347,
        'ra' => 3.6818738679550713,
        'dec' => -1.0537061748654932,
        'az' => 3.698003619125585,
        'alt' => 0.7240401439162253
    );
    $first = $fomalhaut;
    $second = $arcturus;
    $test = $vega;
    $initialTime = $first['time'];
}

$firstTelescopeVector = getTelescopeVector($first);
$firstEquatorialVector = getEquatorialVector($first, $initialTime);

$secondTelescopeVector = getTelescopeVector($second);
$secondEquatorialVector = getEquatorialVector($second, $initialTime);

$thirdTelescopeVector = getVectorProduct($firstTelescopeVector, $secondTelescopeVector);
$thirdEquatorialVector = getVectorProduct($firstEquatorialVector, $secondEquatorialVector);

$equatorialMatrix = makeM3($firstEquatorialVector, $secondEquatorialVector, $thirdEquatorialVector);

$inverseM3 = getInverseM3($equatorialMatrix);

$transform = getMatrixProduct(makeM3($firstTelescopeVector, $secondTelescopeVector, $thirdTelescopeVector), $inverseM3);

$inverseTransform = getInverseM3($transform);

/*

echo $first['name']." telescope vector:\n";
printVector($firstTelescopeVector);

echo $first['name']." equatorial vector:\n";
printVector($firstEquatorialVector);

echo $second['name']." telescope vector:\n";
printVector($secondTelescopeVector);

echo $second['name']." equatorial vector:\n";
printVector($secondEquatorialVector);

echo "third telescope vector:\n";
printVector($thirdTelescopeVector);

echo "third equatorial vector:\n";
printVector($thirdEquatorialVector);

echo "Equatorial matrix:\n";
printM3($equatorialMatrix);

echo "Inverse equatorial matrix:\n";
printM3($inverseM3);

echo "Transform matrix:\n";
printM3($transform);

echo "Inverse transform matrix:\n";
printM3($inverseTransform);

*/

echo $first['name']." telescope coords:\n";
printVectorAsDMS($firstTelescopeVector);

echo $first['name']." equatorial coords:\n";
printVectorAsHMS($firstEquatorialVector);

echo $second['name']." telescope coords:\n";
printVectorAsDMS($secondTelescopeVector);

echo $second['name']." equatorial coords:\n";
printVectorAsHMS($secondEquatorialVector);

echo "third telescope coords:\n";
printVectorAsDMS($thirdTelescopeVector);

echo "third equatorial coords:\n";
printVectorAsHMS($thirdEquatorialVector);

echo "third telescope vector:\n";
printVector($thirdTelescopeVector);

echo "third equatorial vector:\n";
printVector($thirdEquatorialVector);

$testEquatorialVector = getEquatorialVector($test, $initialTime);

echo $test['name']." equatorial vector:\n";
printVector($testEquatorialVector);

$testTelescopeVector = getMatrixProduct($transform, $testEquatorialVector, 3, 3, 1);

echo $test['name']." telescope vector:\n";
printVector($testTelescopeVector);

$testTelescopeCoords = getCoordsFromVector($testTelescopeVector);

echo $test['name']." telescope coordinates calculated from equatorial data:\n";
echo "Alt: ".radToDMS($testTelescopeCoords[1])." Az: ".radToDMS($testTelescopeCoords[0])."\n\n";

echo $test['name']." telescope coordinates calculated from equatorial data in rads:\n";
echo "Alt: ".($testTelescopeCoords[1])." Az: ".($testTelescopeCoords[0]);


// equation 5.4-5
function getTelescopeVector($star) {
    return array(
        cos($star['alt']) * cos($star['az']),
        cos($star['alt']) * sin($star['az']),
        sin($star['alt'])
    );
}

// equation 5.4-6
function getEquatorialVector($star, $initialTime) {
    return array(
        cos($star['dec']) * cos($star['ra'] - $GLOBALS['k'] * ($star['time'] - $initialTime)),
        cos($star['dec']) * sin($star['ra'] - $GLOBALS['k'] * ($star['time'] - $initialTime)),
        sin($star['dec'])
    );
}

function getVectorProduct($v1, $v2) {
    $multiplier = 1 / sqrt(
        pow($v1[1] * $v2[2] - $v1[2] * $v2[1], 2) +
        pow($v1[2] * $v2[0] - $v1[0] * $v2[2], 2) +
        pow($v1[0] * $v2[1] - $v1[1] * $v2[0], 2)
    );
    
    return array(
        $multiplier * ($v1[1] * $v2[2] - $v1[2] * $v2[1]),
        $multiplier * ($v1[2] * $v2[0] - $v1[0] * $v2[2]),
        $multiplier * ($v1[0] * $v2[1] - $v1[1] * $v2[0])
    );
}

function getCoordsFromVector($v) {
    $ha = atan($v[1] / $v[0]) + ($v[0] < 0 ? pi() : 0);
    if ($ha < 0) $ha += 2.0 * pi();
    return array($ha, asin($v[2]));
}

/*
function radToDMS($r) {
    $decimalDegrees = abs($r * 180.0 / pi());
    $d = floor($decimalDegrees);
    $m = floor(($decimalDegrees - $d) * 60.0);
    $s = ($decimalDegrees - $d - ($m / 60.0)) * 3600.0;
    $d *= ($r < 0 ? -1 : 1);
    return sprintf("% 4d*%02d'%04.1f\"", $d, $m, $s);
}*/

function radToDMS($r) {
    $decimalDegrees = abs($r * 180.0 / pi());
    $d = floor($decimalDegrees);
    $m = floor(($decimalDegrees - $d) * 60.0);
    $s = ($decimalDegrees - $d - ($m / 60.0)) * 3600.0;
    $d *= ($r < 0 ? -1 : 1);
    return sprintf("% 4d*%02d'%04.1f\"", $d, $m, $s);
}

function radToHMS($r) {
    $decimalDegrees = abs($r * 12.0 / pi());
    $d = floor($decimalDegrees);
    $m = floor(($decimalDegrees - $d) * 60.0);
    $s = ($decimalDegrees - $d - ($m / 60.0)) * 3600.0;
    $d *= ($r < 0 ? -1 : 1);
    return sprintf("% 3d:%02d:%04.1f", $d, $m, $s);
}

function printVector($v) {
    echo "L = $v[0]\n";
    echo "M = $v[1]\n";
    echo "N = $v[2]\n\n";
}

function printVectorAsDMS($v) {
    $coords = getCoordsFromVector($v);
    echo "Az:  ".radToDMS($coords[0])."\nAlt: ".radToDMS($coords[1])."\n\n";
}

function printVectorAsHMS($v) {
    $coords = getCoordsFromVector($v);
    echo "RA:   ".radToHMS($coords[0])."\nDec: ".radToDMS($coords[1])."\n\n";
}

function printM3($m3) {
    echo "$m3[0] $m3[1] $m3[2] \n";
    echo "$m3[3] $m3[4] $m3[5] \n";
    echo "$m3[6] $m3[7] $m3[8] \n\n";
}


function makeM3($v1, $v2, $v3) {
    return array($v1[0], $v2[0], $v3[0],
                 $v1[1], $v2[1], $v3[1],
                 $v1[2], $v2[2], $v3[2]);
}

function getMatrixProduct($m1, $m2, $m1row = 3, $m1cols = 3, $m2cols = 3) {
    $result = array();
	for ($i = 0; $i < $m1row; $i++)
		for($j = 0; $j < $m2cols; $j++)
		{
			$result[$m2cols * $i + $j] = 0;
			for ($k = 0; $k < $m1cols; $k++)
				$result[$m2cols * $i + $j]= $result[$m2cols * $i + $j] + $m1[$m1cols * $i + $k] * $m2[$m2cols * $k + $j];
		}
	return $result;
}

function getInverseM3($m3) {
    // A = input matrix AND result matrix
    // n = number of rows = number of columns in A (n x n)
    $n = 3; // M3s are always 3x3
    $pivrow;        // keeps track of current pivot row
    $k;$i;$j;        // k: overall index along diagonal; i: row index; j: col index
    $pivrows = array(); // keeps track of rows swaps to undo at end
    $tmp;        // used for finding max value and making column swaps

    for ($k = 0; $k < $n; $k++)
    {
        // find pivot row, the row with biggest entry in current column
        $tmp = 0;
        for ($i = $k; $i < $n; $i++)
        {
            if (abs($m3[$i * $n + $k]) >= $tmp)    // 'Avoid using other functions inside abs()?'
            {
                $tmp = abs($m3[$i * $n + $k]);
                $pivrow = $i;
            }
        }

        // check for singular matrix
        if ($m3[$pivrow * $n + $k] == 0.0)
        {
            return false;
        }

        // Execute pivot (row swap) if needed
        if ($pivrow != $k)
        {
            // swap row k with pivrow
            for ($j = 0; $j < $n; $j++)
            {
                $tmp = $m3[$k * $n + $j];
                $m3[$k * $n + $j] = $m3[$pivrow * $n + $j];
                $m3[$pivrow * $n + $j] = $tmp;
            }
        }
        $pivrows[$k] = $pivrow;    // record row swap (even if no swap happened)

        $tmp = 1.0 / $m3[$k * $n + $k];    // invert pivot element
        $m3[$k * $n + $k] = 1.0;        // This element of input matrix becomes result matrix

        // Perform row reduction (divide every element by pivot)
        for ($j = 0; $j < $n; $j++)
        {
            $m3[$k * $n + $j] = $m3[$k * $n + $j] * $tmp;
        }

        // Now eliminate all other entries in this column
        for ($i = 0; $i < $n; $i++)
        {
            if ($i != $k)
            {
                $tmp = $m3[$i * $n + $k];
                $m3[$i * $n + $k] = 0.0;  // The other place where in matrix becomes result mat
                for ($j = 0; $j < $n; $j++)
                {
                    $m3[$i * $n + $j] = $m3[$i * $n + $j] - $m3[$k * $n + $j] * $tmp;
                }
            }
        }
    }
    
    for ($k = $n - 1; $k >= 0; $k--)
    {
        if ($pivrows[$k] != $k)
        {
            for ($i = 0; $i < $n; $i++)
            {
                $tmp = $m3[$i * $n + $k];
                $m3[$i * $n + $k] = $m3[$i * $n + $pivrows[$k]];
                $m3[$i * $n + $pivrows[$k]] = $tmp;
            }
        }
    }
    
    return $m3;
}

