/*****************************************************************************\
 * RiseSet.cpp  (meh)
 *
 * Calculates rise & set times of sun or moon, plus twilight times
 *
 * author: mark huss (mark@mhuss.com)
 * Based on Bill Gray's open-source code at projectpluto.com
 *
\*****************************************************************************/

#include "RiseSet.h"
#include "PlanetData.h"
#include "Vsop.h"

#include <math.h>

// Altitudes for each event
//
static const double SUN_ALT = Astro::toRadians(-.83333);
static const double MOON_ALT = Astro::toRadians(.125);
static const double C_TWI_ALT = Astro::toRadians(-6.);
static const double N_TWI_ALT = Astro::toRadians(-12.);
static const double A_TWI_ALT = Astro::toRadians(-18.);

static PlanetData pd;

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  This class computes the times at which the sun or moon will rise and set
  during a given day starting on 'jd'.  It does this by computing the altitude
  of the object during each of the 24 hours of that day (and 1 hour of the
  next). What we really want to know is the object's altitude relative to the
  'rise/set altitude' (the altitude at which the top of the object becomes
  visible,  after correcting for refraction and, in the case of the Moon,
  topocentric parallax.)

  For the sun,  this altitude is -.8333 degrees (its apparent radius
  is about .25 degrees,  and refraction 'lifts it up' by .58333 degrees.)
  For the moon,  this altitude is +.125 degrees.

  If we find that the object was below this altitude at one hour,
  and above it on the next hour, then it must have risen in that interval;
  Conversely, if we find that the object was above this altitude at one hour,
  and below it on the next hour, then it must have set in that interval.

  We then do an iterative search to find the instant during that hour
  that it rose or set. This starts with a guessed rise/set time in the middle
  of the particular hour in question. At each step, we look at the altitude of
  that object at that time, and use it to adjust the rise/set time based on
  the assumption that the motion was linear during the hour (this isn't a
  perfect assumption, but we still usually converge in a few iterations.)

  As a side benefit, we this function will also calculate twilight times
  by using the sun and just changing the event altitude. The modified
  altitudes are -6 degrees (civil twilight), -12 degrees (nautical twilight)
  and -18 degrees (astronomical twilight).

  The rise (or twilight start) time is stored in riseSet.a
  The set (or twilight end) time is stored in riseSet.b
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
void RiseSet::getTimes( TimePair& riseSet, RSType rst, double jd, ObsInfo& oi )
{
  double risesetAlt;  // r/s altitude
  Planet planet = EARTH;

  switch ( rst ) {
  case SUN:
    risesetAlt = SUN_ALT;
    break;
  case MOON:
    risesetAlt = MOON_ALT;
    planet = LUNA;    // moon
    break;
  case CIVIL_TWI:
    risesetAlt = C_TWI_ALT;
    break;
  case NAUTICAL_TWI:
    risesetAlt = N_TWI_ALT;
    break;
  case ASTRONOMICAL_TWI:
    risesetAlt = A_TWI_ALT;
    break;
  default:
     break;
  };

  /*
   * Mark both the rise and set times as -1,  to indicate that they've
   * not been found.  Note that it may turn out that one or both do
   * not occur during the given 24 hours.
   */

  riseSet.a = riseSet.b = -1.;

  double altitude[Astro::IHOURS_PER_DAY+1];    // 24 hrs + 1

  // Compute the altitude for each hour:
  //
  for( int i=0; i<=Astro::IHOURS_PER_DAY; i++ ) {
    pd.calc( planet, jd + Astro::toDays(i), oi );
    altitude[i] = asin( pd.altazLoc(2) ) - risesetAlt;
  }

  // Scan the hours, looking for rise/sets:
  //
  for( int i=0; i<Astro::IHOURS_PER_DAY; i++ ) {
    double* pRS = 0;
    if( altitude[i] <= 0. && altitude[i+1] > 0.) {
      // object is rising
      pRS = &riseSet.a;
    }
    else if( altitude[i] > 0. && altitude[i+1] <= 0. ) {
      // object is setting
      pRS = &riseSet.b;
    }

    if ( 0 != pRS ) {
      // we found a rise or set to refine

      double fraction = Astro::toDays(i);
      double alt0 = altitude[i];
      double altDiff = altitude[i+1] - alt0;
      double delta = 1.;
      int iterations = 10;

      while( delta > .0001 && iterations-- ) {
        delta = ( -alt0 / altDiff ) / Astro::HOURS_PER_DAY;
        fraction += delta;
        pd.calc( planet, jd + fraction, oi );
        alt0 = asin( pd.altazLoc(2) ) - risesetAlt;
      }
      *pRS = fraction;
    }
  }
}

