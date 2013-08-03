#include <stdio.h>
#include <stdlib.h>
#include "model.h"

#define PI 3.1415926535897932384626433
#define MAX_STARS 300

/* Cheap and simple demonstration of how the scope modelling works */
/* Just compile as cl -W4 -Ox mod_test.cpp model.cpp               */

void main( int argc, char **argv)
{
   FILE *ifile = fopen( "align.out", "rb");
   double *model_params, x, y, alt, az;
   char buff[80];
   int n_params, i;

   if( !ifile)
      {
      printf( "ALIGN.OUT not found\n");
      exit( 0);
      }

   fgets( buff, sizeof( buff), ifile);       /* skip this line... */
   fgets( buff, sizeof( buff), ifile);
   n_params = atoi( buff);
   model_params = (double *)calloc( n_params, sizeof( double));
   for( i = 0; i < n_params; i++)
      {
      fgets( buff, sizeof( buff), ifile);
      model_params[i] = atof( buff);
      }
   fclose( ifile);
   printf( "Got the model loaded from ALIGN.OUT\n");
   printf( "Enter a pair of encoder coordinates: ");
   scanf( "%lf %lf", &x, &y);
   encoder_to_alt_az( model_params, x, y, &alt, &az);
   while( az > PI + PI)
      az -= PI + PI;
   while( az < 0.)
      az += PI + PI;
   printf( "In alt/az values: alt %.4lf, az %.4lf\n",
      alt * 180. / PI, az * 180. / PI);
   alt_az_to_encoder( model_params, alt, az, &x, &y);
   printf( "Recovered x/y: %lf %lf\n", x, y);
}
