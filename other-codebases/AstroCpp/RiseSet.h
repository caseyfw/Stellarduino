/*****************************************************************************\
 * RiseSet.h
 *
 * Calculates rise & set times of sun or moon
 * Also calculates twilight times (civil/nautical/astronomical)
 *
 * author: mark huss (mark@mhuss.com)
 * Based on Bill Gray's open-source code at projectpluto.com
 *
\*****************************************************************************/

#if !defined( RISE_SET__H )
#define RISE_SET__H

#include "AstroOps.h"

// A pair of times
struct TimePair {
   double a, b;
};

// A pair of times, specificly rise & set times
struct TimePairRS {
   double rise, set;
};

class ObsInfo;

class RiseSet {
public:

  // The types of time pair events we support

  enum RSType {
      SUN = 0, MOON = 1,
      CIVIL_TWI = 2, NAUTICAL_TWI = 3, ASTRONOMICAL_TWI = 4
  };

  /*
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
  */
  static void getTimes( TimePair& riseSet, RSType rst, double jd, ObsInfo& oi);

  /*
  The 'quadrant' function helps in figuring out dates of lunar phases
  and solstices/equinoxes.  If the solar longitude is in one quadrant at
  the start of a day,  but in a different quadrant at the end of a day,
  then we know that there must have been a solstice or equinox during that
  day.  Also,  if (lunar longitude - solar longitude) changes quadrants
  between the start of a day and the end of a day,  we know there must have
  been a lunar phase change during that day.
  */

  static int quadrant( double angle ) {
    return (int)( AstroOps::normalizeRadians( angle ) * Astro::TWO_OVER_PI );
  }

};

#endif  /* #if !defined( RISE_SET__H ) */
