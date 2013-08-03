#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <math.h>
#include <time.h>
#include "watdefs.h"
#include "afuncs.h"
#include "lunar.h"

/* The following strips comments and extra spaces out of 'cospar.txt'.
I thought this might produce a faster version.  It doesn't;  the
COSPAR routines do not appear to be greatly troubled by the extra
stuff in 'cospar.txt'.  Most of the time appears to be spent in
doing math,  as I'd pretty much expected.          */

void produce_fixed_cospar( void)
{
   FILE *ifile = fopen( "cospar.txt", "rb");
   char buff[400];
   int i;

   while( fgets( buff, sizeof( buff), ifile) && memcmp( buff, "END", 3))
      if( *buff >= ' ' && *buff != '#')
         {
         char *tptr = strchr( buff, '(');

         if( tptr)
            *tptr = '\0';
         for( i = 0; buff[i]; i++)
            if( buff[i] == ' ' && buff[i + 1] == ' ')
               {
               memmove( buff + i, buff + i + 1, strlen( buff + i));
               i--;
               }
         for( i = 0; buff[i] >= ' '; i++)
            ;
         buff[i] = '\0';
         printf( "%s\n", buff);
         }
   fclose( ifile);
}

      /* Unit test code for COSPAR functions.  I've used this     */
      /* after making changes to 'cospar.txt' or 'cospar.cpp'     */
      /* just to verify that only the things that were _supposed_ */
      /* to change,  actually changed.                            */

int main( int argc, char **unused_argv)
{
   double matrix[9], prev_matrix[9];
   int i, j, system_number, rval;
   clock_t t0 = clock( );

   setvbuf( stdout, NULL, _IONBF, 0);
   if( argc == 2)
      produce_fixed_cospar( );
   for( i = 0; i < 9; i++)
      prev_matrix[i] = 0.;
   for( i = 0; i < 2000; i++)
      {
      for( system_number = 0; system_number < 4; system_number++)
         {
         rval = calc_planet_orientation( i, system_number,
                        2451000. + (double)i * 1000., matrix);
         if( rval && rval != -1)
            printf( "rval %d\n", rval);
         else if( !rval)
            if( memcmp( matrix, prev_matrix, 9 * sizeof( double)))
               {
               printf( "Planet %d, system %d\n", i, system_number);
               for( j = 0; j < 9; j += 3)
                  printf( "%11.8lf %11.8lf %11.8lf\n",
                             matrix[j], matrix[j + 1], matrix[j + 2]);
               for( j = 0; j < 9; j++)
                  prev_matrix[j] = matrix[j];
               }
         }
      }
   printf( "Total time: %lf\n",
            (double)( clock() - t0) / (double)CLOCKS_PER_SEC);
   return( 0);
}
