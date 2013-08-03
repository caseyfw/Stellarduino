#include <math.h>
#include "model.h"

static void make_vect( double *vect, const double az, const double alt)
{
   double cos_alt = cos( alt);

   *vect++ = cos( az) * cos_alt;
   *vect++ = sin( az) * cos_alt;
   *vect++ = sin( alt);
}

int encoder_to_alt_az( const double *m, const double x, const double y,
                           double *alt, double *az)
{
   const double pi = 3.141592653589793238462643383279;
   double pseudo_az  = m[MODEL_AZ_ZERO]  + m[MODEL_AZ_RATE] * x;
   double pseudo_alt = m[MODEL_ALT_ZERO] + m[MODEL_ALT_RATE] * y;
   double vect[3], real_alt, delta_az, real_az;
   const double cos_misalign = cos( m[MODEL_MISALIGN]);
// double tan_misalign = sin( m[MODEL_MISALIGN]) / cos_misalign;
   int i;

   pseudo_alt += m[MODEL_FLEXURE] * cos( pseudo_alt) +
                 m[MODEL_WARPED_TBL_1] * cos( pseudo_az * 2.) +
                 m[MODEL_WARPED_TBL_2] * sin( pseudo_az * 2.) +
                 m[MODEL_ALT_SIN] * sin( pseudo_alt) +
                 m[MODEL_ALT_SIN2] * sin( pseudo_alt * 2.) +
                 m[MODEL_ALT_COS2] * cos( pseudo_alt * 2.);
   pseudo_az += m[MODEL_AZ_COS] * cos( pseudo_az) +
                m[MODEL_AZ_SIN] * sin( pseudo_az) +
                m[MODEL_AZ_SIN2] * sin( pseudo_az * 2.) +
                m[MODEL_AZ_COS2] * cos( pseudo_az * 2.);
               /* now correct for misalign,  the non-perpendicularity of the */
               /* alt and az axes: */

   if( !m[MODEL_MISALIGN])
      {
      real_alt = pseudo_alt;
      delta_az = 0.;
      }
   else
      {
      real_alt = asin( sin( pseudo_alt) * cos_misalign);
      delta_az = acos( cos( pseudo_alt) / cos( real_alt));
      if( m[MODEL_MISALIGN] < 0.)
         delta_az = -delta_az;

                  /* 2 Jun 1999:  the following line attempts to handle */
                  /* some 'over the pole' weirdness.  If pseudo_alt winds */
                  /* up on the opposite side of the pole,  we have to     */
                  /* adjust the azimuth accordingly. */

//    if( cos( pseudo_alt) < 0.)
//       delta_az = pi + delta_az;
      }

   real_az = pseudo_az + delta_az;

               /* and now go to vector form,  so we can tilt it twice: */
   make_vect( vect, real_az, real_alt);

   for( i = 0; i < 2; i++)
      {
      double tilt = m[MODEL_TILT1 + i];

      double temp = sin( tilt) * vect[i] + cos( tilt) * vect[2];
      vect[i] = cos( tilt) * vect[i] - sin( tilt) * vect[2];
      vect[2] = temp;
      }
   *az = atan2( vect[1], vect[0]);
   *alt = asin( vect[2]);
   return( 0);
}

int alt_az_to_encoder( const double *m, const double alt, const double az,
                           double *x, double *y)
{
   double pseudo_az, pseudo_alt;
   double vect[3];
   double cos_misalign = cos( m[MODEL_MISALIGN]);
   double tan_misalign = sin( m[MODEL_MISALIGN]) / cos_misalign;
   int i;

   make_vect( vect, az, alt);

   for( i = 1; i >= 0; i--)
      {
      double tilt = m[MODEL_TILT1 + i];

      double temp = -sin( tilt) * vect[i] + cos( tilt) * vect[2];
      vect[i] = cos( tilt) * vect[i] + sin( tilt) * vect[2];
      vect[2] = temp;
      }

   pseudo_az = atan2( vect[1], vect[0]);
   pseudo_alt = asin( vect[2]);
                                    /* now correct for non-perpen. axes: */
   pseudo_az -= vect[2] * tan_misalign / cos( pseudo_alt);
   pseudo_alt = asin( sin( pseudo_alt) / cos_misalign);

   pseudo_alt -= m[MODEL_FLEXURE] * cos( pseudo_alt) +
                 m[MODEL_WARPED_TBL_1] * cos( pseudo_az * 2.) +
                 m[MODEL_WARPED_TBL_2] * sin( pseudo_az * 2.) +
                 m[MODEL_ALT_SIN] * sin( pseudo_alt) +
                 m[MODEL_ALT_SIN2] * sin( pseudo_alt * 2.) +
                 m[MODEL_ALT_COS2] * cos( pseudo_alt * 2.);
   pseudo_az -= m[MODEL_AZ_COS] * cos( pseudo_az) +
                m[MODEL_AZ_SIN] * sin( pseudo_az) +
                m[MODEL_AZ_SIN2] * sin( pseudo_az * 2.) +
                m[MODEL_AZ_COS2] * cos( pseudo_az * 2.);

   *x = (pseudo_az - m[MODEL_AZ_ZERO]) / m[MODEL_AZ_RATE];
   *y = (pseudo_alt - m[MODEL_ALT_ZERO]) / m[MODEL_ALT_RATE];
   return( 0);
}
