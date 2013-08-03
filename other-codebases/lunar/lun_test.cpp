#include <stdio.h>
#include <stdlib.h>
#include "lun_tran.h"

int main( const int argc, const char **argv)
{
   double lon, lat;
   int year = atoi( argv[1]), month = atoi( argv[2]);
   int day, zip_code = atoi( argv[3]);
   int time_zone, use_dst;
   char place_name[100];

   if( get_zip_code_data( zip_code, &lat, &lon, &time_zone,
                           &use_dst, place_name))
      {
      printf( "Couldn't get data for ZIP code %05d\n", zip_code);
      return( -1);
      }

   printf( "Lat %lf   Lon %lf   Time zone %d  DST=%d  %s\n",
         lat, lon, time_zone, use_dst, place_name);
   for( day = 1; day <= 31; day++)
      {
      const double transit_time =
            get_lunar_transit_time( year, month, day,
                  lat, lon, time_zone, use_dst, 1);
      const double antitransit_time =
            get_lunar_transit_time( year, month, day,
                  lat, lon, time_zone, use_dst, 0);
      char transit_buff[6], antitransit_buff[6];

      format_hh_mm( transit_buff, transit_time);
      format_hh_mm( antitransit_buff, antitransit_time);

      printf( "%2d: %s %s\n", day,
            transit_buff, antitransit_buff);
      }
   return( 0);
}
