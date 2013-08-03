#include <stdio.h>
#include "DateOps.h"
#include "Lunar.h"

int main(int argc, char** argv) {

  double jd = 0;
  double tz = -5./24.;
  for ( int d = 24; d < 25; d++ ) {
    for ( int h = 0; h < 24; h++ ) {
      jd = DateOps::dmyToDay( d, 1, 2001 ) + tz + h/24.;
      printf( "d=%d, h=%d, age=%2.3f\n", d, h, Lunar::ageOfMoonInDays(jd) );
    }
  }
}
