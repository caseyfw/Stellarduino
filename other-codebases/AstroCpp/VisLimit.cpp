//----------------------------------------------------------------------------
// VisLimit - calculates the visual limiting magnitude
//
// Astro library based on open-source code from project pluto
//
// mark huss 11/2000
//----------------------------------------------------------------------------

/*
The computations for sky brightness and limiting magnitude can be
logically broken up into several pieces.  Some computations depend
on things that are constant for a given observing site and time:
the lunar and solar zenith distances,  the air masses to those objects,
the temperature and relative humidity,  and so forth.  For use in Guide,
I expect to compute brightness at many points in the sky,  while all
these other values hold constant.  So my first step (after putting
lat/lon and these other data into the BRIGHTNESS_DATA struct) is to
call the set_brightness_params() function.  This function does a lot
of "setup work",  figuring out the absorption per unit air mass at
various wavelengths from various causes (gas,  aerosol,  ozone),
the number of air masses to the sun and moon,  and so forth.

Once you've done all this,  you can call compute_sky_brightness() for
any point in the sky.  You do need to provide the zenith angle,  and the
angular distance of that point from the moon and sun.  The brightnesses
are returned in the brightness[] array.  The 'mask' value can be used to
specify which of the five bands is to be computed.  (For example,  if I
use this to make a realistic sky background,  I may just concern myself
with the V band... maybe with B and R if I want to attempt a colored sky.
In either case,  computing all five bands would be excessive.)

Next,  you can call compute_extinction( ) to set any or all of the five
extinction values.  Normally,  I wouldn't see much use for this data.
But you do need to have that data if you intend to call the
compute_limiting_mag( ) function.

All of what follows is adapted from Brad Schaefer's article and code
on pages 57-60,  May 1998 _Sky & Telescope_,  "To the Visual Limits".

NOTICE that I modified his test conditions.  He had the moon and sun
well below the horizon;  I found that this didn't make testing
contributions from those objects any easier,  so I put them where they
could contribute more brightness.

At some point when I have the time,  I'll break out the main( ) portion,
tack in the code for CCD mag limits on page 121 of the same magazine,
and make proper header files.

*/

#include <math.h>
#include <stdlib.h>
#include "VisLimit.h"

#define MAG_TO_BRIGHTNESS( X) (exp( -.4 * (X) * LOG_10))
#define BRIGHTNESS_TO_MAG( X) (-2.5 * log( X) / LOG_10)
#define PI 3.141592653589793
#define LOG_10 2.302585093

double VisLimit::computeAirMass( const double zenithAngle)
{
   double rval = 40., cosAng = cos( zenithAngle);

   if( cosAng > 0.)
      rval = 1. / (cosAng + .025 * exp( -11. * cosAng));

   return( rval);
}

double VisLimit::computeFFactor( double objDist)
{
   double objDistDegrees = objDist * 180. / PI;
   double rval, cosDist = cos( objDist);

   rval = 6.2e+7 / (objDistDegrees * objDistDegrees)
                        + exp( LOG_10 * (6.15 - objDistDegrees / 40.));
   rval += 229086. * (1.06 + cosDist * cosDist);  /* polarization term? */
   return( rval);
            /* Seen on lines 2210 & 2200 for the moon,  and on lines */
            /* 2320 & 2330 for the moon.  I've only foggy ideas what  */
            /* it means;  I think it attempts to compute the falloff in */
            /* scattered light from an object as a function of distance.  */
}

int VisLimit::setBrightnessParams(FixedBrightnessData& b)
{
  fixed = b;
  double monthAngle = (fixed.month - 3.) * PI / 6.;
  double kaCoeff, krCoeff, koCoeff, kwCoeff, moonElong;
  int i;

  krCoeff = .1066 * exp( -fixed.htAboveSeaInMeters / 8200.);
  kaCoeff = .1 * exp( -fixed.htAboveSeaInMeters / 1500.);
  if( fixed.relativeHumidity > 0.)
    {
    double humidityParam;

    if( fixed.relativeHumidity >= 100.)
        humidityParam = 1000000.;
    else
        humidityParam = 1. - .32 / log( fixed.relativeHumidity / 100.);
    kaCoeff *= exp( 1.33 * log( humidityParam));
    }
  if( fixed.latitude < 0.)
    kaCoeff *= 1. - sin( monthAngle);
  else
    kaCoeff *= 1. + sin( monthAngle);
  koCoeff = (3. + .4 * (fixed.latitude * cos( monthAngle) -
                    cos( 3. * fixed.latitude))) / 3.;
  kwCoeff = .94 * (fixed.relativeHumidity / 100.) *
                      exp( fixed.temperatureInC / 15.) *
                      exp( -fixed.htAboveSeaInMeters / 8200.);

  yearTerm = 1. + .3 * cos( 2. * PI * (fixed.year - 1992) / 11.);
  airMassMoon = computeAirMass( fixed.zenithAngMoon);
  airMassSun  = computeAirMass( fixed.zenithAngSun);
  moonElong = fixed.moonElongation * 180. / PI;
  lunarMag = -12.73 + moonElong * (.026 +
                          4.e-9 * (moonElong * moonElong * moonElong));
              /* line 2180 in B Schaefer code */
  for( i = 0; i < 5; i++)
    {
    static const double fourthPowerTerms[5] =
                      { 5.155601, 2.441406, 1., 0.381117, 0.139470 };
    static const double onePointThreePowerTerms[5] =
                      { 1.704083, 1.336543, 1., 0.730877, 0.527177 };
    static const double oz[5] = {0., 0., .031, .008, 0.};
    static const double wt[5] = {.074, .045, .031, .02, .015};

    kr[i] = krCoeff * fourthPowerTerms[i];
    ka[i] = kaCoeff * onePointThreePowerTerms[i];
    ko[i] = koCoeff * oz[i];
    kw[i] = kwCoeff * wt[i];

    k[i] = kr[i] + ka[i] + ko[i] + kw[i];
    c3[i] = MAG_TO_BRIGHTNESS( k[i] * airMassMoon);
            /* compute dropoff in lunar brightness from extinction: 2200 */
    c4[i] = MAG_TO_BRIGHTNESS( k[i] * airMassSun);
    }
  return( 0);
}


/* If all you want is the sky brightness,  all the data concerning */
/* separate air masses for gas, aerosols,  and ozone and such is   */
/* an unnecessary drain on computation.  So that's broken out as a */
/* separate process in computeExtinction(). */

int VisLimit::computeExtinction()
{
   double cosZenithAng = cos( angular.zenithAngle );
   double tval;
   int i;

   airMassGas =
               1. / (cosZenithAng + .0286 * exp( -10.5 * cosZenithAng));
   airMassAerosol =
               1. / (cosZenithAng + .0123 * exp( -24.5 * cosZenithAng));
   tval = sin( angular.zenithAngle ) / (1. + 20. / 6378.);
   airMassOzone = 1. / sqrt( 1. - tval * tval);
   for( i = 0; i < 5; i++)
      if( (mask >> i) & 1)
         extinction[i] = (kr[i] + kw[i]) * airMassGas +
                             ka[i] * airMassAerosol +
                             ko[i] * airMassOzone;
   return( 0);
}

double VisLimit::computeLimitingMag()
{
   double c1, c2, bl = brightness[2] / 1.11e-15;
   double th, tval, rval;

   if( bl > 1500.) {
     c1 = 4.4668e-9;
     c2 = 1.2589e-6;
   }
   else {
     c1 = 1.5849e-10;
     c2 = 1.2589e-2;
   }
   tval = 1. + sqrt( c2 * bl);
   th = c1 * tval * tval;        // brightness in foot-candles?
   rval = -16.57 + BRIGHTNESS_TO_MAG( th) - extinction[2];
   return( rval);
}

int VisLimit::computeSkyBrightness(AngularBrightnessData& abd)
{
   angular = abd;

   double sinZenith;
   double brightnessDrop2150, fs, fm;
   int i;

   double airMass = computeAirMass( angular.zenithAngle );
   sinZenith = sin( angular.zenithAngle );
   brightnessDrop2150 = .4 + .6 / sqrt( 1.0 - .96 * sinZenith * sinZenith);
   fm = computeFFactor( angular.distMoon );
   fs = computeFFactor( angular.distSun );

   for( i = 0; i < 5; i++)
      if( (mask >> i) & 1)
         {
         static const double bo[5] = {8.0e-14, 7.e-14, 1.e-13, 1.e-13, 3.e-13};
               /* Base sky brightness in each band */
         static const double cm[5] = {1.36, 0.91, 0.00, -0.76, -1.17 };
               /* Correction to moon's magnitude */
         static const double ms[5] = {-25.96, -26.09, -26.74, -27.26, -27.55 };
               /* Solar magnitude? */
         static const double mo[5] = {-10.93, -10.45, -11.05, -11.90, -12.70 };
               /* Lunar magnitude? */
         double bn = bo[i] * yearTerm, directLoss;
                        /* accounts for a 30% variation due to sunspots? */

         double brightnessMoon, twilightBrightness;
         double brightnessDaylight;

         directLoss = MAG_TO_BRIGHTNESS( k[i] * airMass);
         bn *= brightnessDrop2150;
                    /* Not sure what this is.. line 2150 in B Schaefer code */
         bn *= directLoss;
                   /* drop brightness to account for extinction: 2160 */

         if( fixed.zenithAngMoon < PI / 2.)      /* moon is above horizon */
            {
            brightnessMoon = MAG_TO_BRIGHTNESS( lunarMag + cm[i]
                                 - mo[i] + 43.27);
            brightnessMoon *= (1. - directLoss);
                  /* Maybe computing how much of the lunar light gets */
                  /* scattered?   2240 */
            brightnessMoon *= (fm * c3[i] + 440000. * (1. - c3[i]));
            }
         else
            brightnessMoon = 0.;

         twilightBrightness = ms[i] - mo[i] + 32.5 -
                           (90. - fixed.zenithAngSun * 180. / PI) -
                           angular.zenithAngle / (2 * PI * k[i]);
                  /* above is in magnitudes,  so gotta do this: */
         twilightBrightness = MAG_TO_BRIGHTNESS( twilightBrightness);
                  /* above is line 2280,  B Schaefer code */
         twilightBrightness *= 100. / (angular.distSun * 180. / PI);
         twilightBrightness *= 1. - MAG_TO_BRIGHTNESS( k[i]);
                  /* preceding line looks suspicious to me... line 2290 */
         brightnessDaylight = MAG_TO_BRIGHTNESS( ms[i] - mo[i] + 43.27);
                     /* line 2340 */
         brightnessDaylight *= (1. - directLoss);
                     /* line 2350 */
         brightnessDaylight *= fs * c4[i] + 440000. * (1. - c4[i]);
         if( brightnessDaylight > twilightBrightness)
            brightness[i] = bn + twilightBrightness + brightnessMoon;
         else
            brightness[i] = bn + brightnessDaylight + brightnessMoon;
#ifdef TEST_STATEMENTS
         if( i == 0)
            printf( "Brightnesses: %lg %lg %lg %lg\n", bn,
                  brightnessMoon, twilightBrightness, brightnessDaylight);
#endif
         }
   return( 0);
}

#ifdef TEST_PROGRAM
#include <stdio.h>

int main( int argc, char **argv)
{
   if (argc < 2) {
     fprintf( stderr, "usage: %s <zenith angle>\n", argv[0] );
     return 1;
   }
   FixedBrightnessData f;
   AngularBrightnessData a;
   VisLimit v;
   int i;

   f.zenithAngMoon = 40. * PI / 180.;
   f.zenithAngSun = 100. * PI / 180.;
   f.moonElongation = 180. * PI / 180.;        // full moon
   f.htAboveSeaInMeters = 1000.;
   f.latitude = 30. * PI / 180.;
   f.temperatureInC = 15.;
   f.relativeHumidity = 40.;
   f.year = 2000.;
   f.month = 11.;
   v.setBrightnessParams(f);

   // values varying across the sky:
   a.zenithAngle = atof( argv[1]) * PI / 180.;
   a.distMoon = 50. * PI / 180.;
   a.distSun = 40. * PI / 180.;

   v.setMask(31);

   v.computeSkyBrightness(a);
   v.computeExtinction();
   for( i = 0; i < 5; i++)
      printf( "%lf  %lg  %.5lf\n",
          v.getK(i), v.getBrightness(i), v.getExtinction(i));
   printf( "Limiting magnitude: %.5lf", v.computeLimitingMag());

   return 0;
}
#endif
