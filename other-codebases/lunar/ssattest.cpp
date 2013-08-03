#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include "watdefs.h"
#include "lunar.h"

int main( const int argc, const char **argv)
{
   int i;
   double loc[3], t;

   t = atof( argv[1]);
   printf( "Date: %.5lf\n", t);
   for( i = 0; i < 8; i++)
      {
      t = atof( argv[1]);
      calc_ssat_loc( t, loc, i, 0L);
      printf( "%d: %9.6lf %9.6lf %9.6lf\n", i, loc[0] * 100., loc[1] * 100.,
                  loc[2] * 100.);
      }
   return( 0);
}
