/* classel.cpp: converts state vects to classical elements

Copyright (C) 2010, Project Pluto

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
02110-1301, USA.    */

#include <math.h>
#include "watdefs.h"
#include "afuncs.h"
#include "comets.h"

#define PI 3.1415926535897932384626433832795028841971693993751058209749445923
#define SQRT_2 1.41421356

/* 2011 Aug 11:  in dealing with exactly circular orbits,  and those
   nearly exactly circular,  I found several loss of precision problems
   in the angular elements (which become degenerate for e=0;  you can
   add an arbitrary amount to the mean anomaly,  as long as you subtract
   the same amount from the argument of periapsis.)  Also,  q was
   computed incorrectly for such cases when roundoff error resulted in
   the square root of a number that should be zero,  but rounded below
   zero,  was taken.  All this is fixed now.       */

/* 2009 Nov 24:  noticed a loss of precision problem in computing arg_per.
   This was done by computing the cosine of that value,  then taking the
   arc-cosine.  But if that value is close to +/-1,  precision is lost
   (you can actually end up with a domain error if the roundoff goes
   against you).  I added code so that,  if |cos_arg_per| > .7,  we
   compute the _sine_ of the argument of periapsis and use that instead.

   While doing this,  I also noticed that several variables could be made
   of type const.   */

/* calc_classical_elements( ) will take a given state vector r at a time t,
   for an object orbiting a mass gm;  and will compute the orbital elements
   and store them in the elem structure.  Normally,  ref=1.  You can set
   it to 0 if you don't care about the angular elements (inclination,
   longitude of ascending node,  argument of perihelion).         */

static double dot_product( const double *a, const double *b)
{
   return( a[0] * b[0] + a[1] * b[1] + a[2] * b[2]);
}

/* MSVC lacks inverse hyperbolic functions: */

#ifdef _MSC_VER
static double atanh( const double x)
{
   return( .5 * log( (1. + x) / (1. - x)));
}
#endif

int DLL_FUNC calc_classical_elements( ELEMENTS *elem, const double *r,
                             const double t, const int ref, const double gm)
{
   const double *v = r + 3;
   const double r_dot_v = dot_product( r, v);
   const double dist = vector3_length( r);
   const double v2 = dot_product( v, v);
   const double inv_major_axis = 2. / dist - v2 / gm;
   double h0, n0, tval;
   double h[3], e[3], ecc2;
   double ecc;
   int i;

   vector_cross_product( h, r, v);
   n0 = h[0] * h[0] + h[1] * h[1];
   h0 = n0 + h[2] * h[2];
   n0 = sqrt( n0);
   h0 = sqrt( h0);

                        /* See Danby,  p 204-206,  for much of this: */
   if( ref & 1)
      {
      elem->asc_node = atan2( h[0], -h[1]);
      elem->incl = asine( n0 / h0);
      if( h[2] < 0.)                   /* retrograde orbit */
         elem->incl = PI - elem->incl;
      }
   vector_cross_product( e, v, h);
   for( i = 0; i < 3; i++)
      e[i] = e[i] / gm - r[i] / dist;
   tval = dot_product( e, h) / h0;     /* "flatten" e vector into the rv */
   for( i = 0; i < 3; i++)             /* plane to avoid roundoff; see   */
      e[i] -= h[i] * tval;             /* above comments                 */
   ecc2 = dot_product( e, e);
   elem->minor_to_major = sqrt( fabs( 1. - ecc2));
   ecc = elem->ecc = sqrt( ecc2);

   if( !ecc)                     /* for purely circular orbits,  e is */
      {                          /* arbitrary in the orbit plane; choose */
      for( i = 0; i < 3; i++)    /* r normalized                         */
         e[i] = r[i] / dist;
      }
   else                           /* ...and if it's not circular,  */
      for( i = 0; i < 3; i++)     /* normalize e:                  */
         e[i] /= ecc;
   if( inv_major_axis)
      {
      elem->major_axis = 1. / inv_major_axis;
      elem->t0 = elem->major_axis * sqrt( fabs( elem->major_axis) / gm);
      }

   if( ecc < .9)
      elem->q = elem->major_axis * (1. - ecc);
   else        /* at eccentricities near one,  the above suffers  */
      {        /* a loss of precision problem,  and we switch to: */
      const double gm_over_h0 = gm / h0;
      const double perihelion_speed = gm_over_h0 +
                   sqrt( gm_over_h0 * gm_over_h0 - inv_major_axis * gm);

      elem->q = h0 / perihelion_speed;
      }

   vector_cross_product( elem->sideways, h, e);
         /* At this point,  elem->sideways has length h0.  */
   if( ref & 1)
      {
      const double cos_arg_per = (h[0] * e[1] - h[1] * e[0]) / n0;

      if( cos_arg_per < .7 && cos_arg_per > -.7)
         elem->arg_per = acos( cos_arg_per);
      else
         {
         const double sin_arg_per =
               (e[0] * h[0] * h[2] + e[1] * h[1] * h[2] - e[2] * n0 * n0)
                                            / (n0 * h0);

         elem->arg_per = fabs( asin( sin_arg_per));
         if( cos_arg_per < 0.)
            elem->arg_per = PI - elem->arg_per;
         }
      if( e[2] < 0.)
         elem->arg_per = PI + PI - elem->arg_per;
      }

   if( inv_major_axis)
      {
      const double r_cos_true_anom = dot_product( r, e);
      const double r_sin_true_anom = dot_product( r, elem->sideways) / h0;
      const double cos_E = r_cos_true_anom * inv_major_axis + ecc;
      const double sin_E = r_sin_true_anom * inv_major_axis
                                        / elem->minor_to_major;

      if( inv_major_axis > 0.)          /* parabolic case */
         {
         const double ecc_anom = atan2( sin_E, cos_E);

         elem->mean_anomaly = ecc_anom - ecc * sin( ecc_anom);
         elem->perih_time = t - elem->mean_anomaly * elem->t0;
         }
      else                             /* hyperbolic case */
         {
         const double ecc_anom = atanh( sin_E / cos_E);

         elem->mean_anomaly = ecc_anom - ecc * sinh( ecc_anom);
         elem->perih_time = t - elem->mean_anomaly * fabs( elem->t0);
         h0 = -h0;
         }
      }
   else              /* parabolic case */
      {
      double tau;

      tau = sqrt( dist / elem->q - 1.);
      if( r_dot_v < 0.)
         tau = -tau;
      elem->w0 = (3. / SQRT_2) / (elem->q * sqrt( elem->q / gm));
/*    elem->perih_time = t - tau * (tau * tau / 3. + 1) *                   */
/*                                      elem->q * sqrt( 2. * elem->q / gm); */
      elem->perih_time = t - tau * (tau * tau / 3. + 1) * 3. / elem->w0;
      }

   for( i = 0; i < 3; i++)
      {
      elem->perih_vec[i] = e[i];
      elem->sideways[i] /= h0;
      }
   elem->angular_momentum = h0;
   return( 0);
}
