
#include <math.h>
double range(); 
extern double rad;

double 
sidereal_time(jd)
double jd;
{
   double st, t, x, ut, e;

   /* find time a UTC 0 hours */
   x = .5 + floor(jd -.5);

/* they a few fractions of a second off, I don't know which one is right */
#ifdef YOU_BELIEVE_THE_NAVY

   /* compute mean sidreal time according to NAVY Almanac */
   t = (x - 2451545.0) / 36525;
   st =  .2790572733+ 100.002139 * t + .0000010776 * t * t;
   st = 24 * (st - floor(st));
   
   /* add in UTC with fudge factor */
   ut = (jd -x) * 24 * 1.002737909;
   st += ut; 

   /* compute correction for equation of equinoxes */
   e = range(125.04452 - 1934.13626 * t + .002071 * t * t);
   e = -.00029 * sin(e * rad) / rad; 
   st += e;
   if (st > 24) st -= 24;
#else

   /* compute GAST according to Meeus Book */
   t = (x - 2415020.0) / 36525;
   st =  .276919398 + 100.0021359 * t + .000001075 * t * t;
   st = 24 * (st - floor(st));
   
   /* add in UTC with fudge factor */
   t = (jd -x) * 24 * 1.002737908;
   st += t; 
   if (st > 24) st -= 24;
#endif

   return st;
}
