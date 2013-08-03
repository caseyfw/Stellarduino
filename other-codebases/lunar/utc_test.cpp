#include <stdio.h>
#include <stdint.h>
#include "watdefs.h"
#include "afuncs.h"

#ifdef _MSC_VER
#define UCONST(a) (a##ui64)
#else
#define UCONST(a) (a##ULL)
#endif

/* Test routine for the utc_minus_ut function in delta_t.cpp.  This function
is a little strange.  To test it out,  the code generates pseudo-random
times between JD 2436000 = 1957 Jun 10.5 and JD 2457000 = 2014 Dec 8.5 and
shows td_minus_utc() for that date.  If you generate output from this program,
modify the utc_minus_ut function,  recompile and re-run,  and get the same
output,  you probably didn't break anything.  Note that the output _will_
change if new leap seconds are added at "unexpected" times.  The code
currently (as of 2013) predicts that the next leap second will be added
at the end of December 2015.  This may prove right,  but it could well be
in June 2015 or June 2016,  or even further off.  You never know what the
earth will do next.

   I'm not using the library rand( ) function,  because I want to be
confident that the results will be the same when compiled on different
systems with different library rand( )s. Following is the MMIX linear
congruential pseudo-random number generator, copied from Donald Knuth.
It's not a great PRNG,  but it's fine for the current humble purpose. */

uint64_t pseudo_random( const uint64_t prev_value)
{
   uint64_t a = UCONST( 6364136223846793005);
   uint64_t c = UCONST( 1442695040888963407);

   return( a * prev_value + c);
}

/* Older MSVC can't handle conversions of unsigned 64-bit integers
   into doubles (!),  but we can get around that as follows: */

#ifdef _MSC_VER
static double uint64_to_double( const uint64_t ival)
{
   const double two_to_the_64th_power = 65536. * 65536. * 65536. * 65536.;
   double rval = (double)( (int64_t)ival);

   if( rval < 0.)
      rval += two_to_the_64th_power;
   return( rval);
}
#endif

int main( const int argc, const char **argv)
{
   uint64_t pseudo = UCONST( 1);
   int i;

   printf( "    JD           TD-UTC   DUT1=UT-UTC\n");

   for( i = 0; i < 10000; i++)
      {
      const double two_to_the_64th_power = 65536. * 65536. * 65536. * 65536.;
#ifdef _MSC_VER
      const double normalized = uint64_to_double( pseudo) / two_to_the_64th_power;
#else
      const double normalized = (double)pseudo / two_to_the_64th_power;
#endif
      const double jd = 2436000. + normalized * (2457000. - 2436000.);

      printf( "%.5lf %11.7lf%11.7lf\n", jd, td_minus_utc( jd),
                  td_minus_ut( jd) - td_minus_utc( jd));
      pseudo = pseudo_random( pseudo);
      }
   return( 0);
}
