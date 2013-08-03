/*----------------------------------------------------------------------------
 * DarkCal - calculates the darkest hours for a given month and year.
 *
 * Created by Mark Huss <mark@mhuss.com>
 *
 * This is a generic console program and has been built and run on both
 * win32 and "unix-like" systems.
 *
 * Developed and built using the mingw32 gcc compiler 2.95.2
 *
 * THIS SOFTWARE IS NOT COPYRIGHTED
 *
 * This source code is offered for use in the public domain. You may
 * use, modify or distribute it freely.
 *
 * This code is distributed in the hope that it will be useful but
 * WITHOUT ANY WARRANTY. ALL WARRANTIES, EXPRESS OR IMPLIED ARE HEREBY
 * DISCLAMED. This includes but is not limited to warranties of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 *
 * Astro library based on Bill Gray's open-source code at projectpluto.com
 *
 * created August 2000
 */
//---------------------------------------------------------------------------

#include "AstroOps.h"
#include "PlanetData.h"
#include "DateOps.h"

#include "RiseSet.h"

#include "ConfigFile.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

//----------------------------------------------------------------------------

#define TP_START a
#define TP_END b
#define TP_RISE a
#define TP_SET b

#define DEBUG 1
#undef PROGRESS_BAR

static const char* CFG_EXT = ".cfg";

static const int DAYS=33;     // 31 max plus one on either side

static const char* monthNames[] = {
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
};

enum DST { DST_NONE, DST_START, DST_END };

// Struct to hold date & location
struct CalData {
  CalData() : month(1), year(2000) {}

  int month;
  int year;
  ObsInfo loc;
};

// file globals
static bool g_tabDelimited = false; // true means produce tab-delimited output
static bool g_ignoreDst = false;    // true means ignore DST
static bool g_html = false;         // true means produce HTML output
static FILE* g_fp = stdout;         // file to use for tab-d or HTML

//----------------------------------------------------------------------------
// print heading
//
void printHeading(CalData& cd) {
  const char* pUtc = ( cd.loc.timeZone() < 0 ) ? "UTC" : "UTC+";
  char title[256];
  sprintf( title,
      "%s, %d at latitude %5.2f, longitude %5.2f, tz %s%d",
       monthNames[ cd.month-1 ], cd.year, cd.loc.degLatitude(),
       cd.loc.degLongitude(), pUtc, cd.loc.timeZone() );

  if ( g_tabDelimited ) {
    fprintf( g_fp, "\n%s\n\n", title );
    fputs( "Day\tDarkest Hours\tEvent\tMoon Rises\tMoon Sets\t"
            "Sunset\tAstronomical Twilight Ends\tNext Day\t"
            "Astronomical Twilight Starts\tSunrise\n", g_fp );
  }
  else if (g_html)
    fprintf( g_fp, "<HTML>\n"
          "<HEAD>\n"
          "  <TITLE>Darkest Hours</TITLE>\n"
          "</HEAD>\n<BODY>\n"
          "<STYLE type=\"text/css\">\n<!--\n"
          "  TH, TD { font-size:12px }\n"
          "  .bar { background:silver }\n"
          "-->\n</STYLE>\n"
          "<H2>Darkest Hours</H2>\n"
          "<TABLE CELLSPACING=0 BORDER=0 WIDTH=\"100%%\">\n"
          "<TR CLASS=\"bar\"><TH COLSPAN=10 ALIGN=\"center\">%s</TH></TR>\n"
          "<TR>\n"
          "  <TH>Day</TH>\n  <TH>Darkest<BR>Hours</TH>\n  <TH>Event</TH>\n"
          "  <TH>Moon<BR>Rises</TH>\n  <TH>Moon<BR>Sets</TH>\n"
          "  <TH>Sunset</TH>\n  <TH>AstTwi<BR>Ends</TH>\n  <TH>Next<BR>Day</TH>\n"
          "  <TH>AstTwi<BR>Starts</TH>\n  <TH>Sunrise</TH>\n"
          "</TR>\n", title );
  else {
    printf( "\n%s\n\n"
            "Day Darkest                  Moon   Moon   Sunset AstTwi Next AstTwi Sunrise\n"
            "    Hours          Events    Rises  Sets          Ends   Day  Starts\n"
            "--  -------------  --------  -----  -----  -----  -----  --   -----  -----\n",
             title );
  }
}

//----------------------------------------------------------------------------
static char* nextColumn( char* p, int sp, bool empty=false )
{
  if (g_tabDelimited)
    *p++ = '\t';
  else if (g_html) {
    if (empty)
      strcpy(p,"&nbsp;</TD>\n  <TD>");
    else
      strcpy(p,"</TD>\n  <TD>");
    p += strlen(p);
  }
  else {
    while (sp--)
      *p++ = ' ';
  }
  return p;
}

//----------------------------------------------------------------------------
// print a (double) time as hh:mm
//
// a time < 0 is printed as '--:--'
//
static char* printTime( char* p, double t, bool ws=true )
{
   if( t < 0.)
      sprintf( p, "--:--" );
   else {
      // round up to nearest minute
      long minutes = long(t * 24. * 60. + .5);
      sprintf( p, "%02d:%02d", int(minutes / 60), int(minutes % 60) );
   }
   return ws ? nextColumn( p+5, 2) : p+5;
}

//----------------------------------------------------------------------------
// print a pair of (double) times as hh:mm
//
static char* printTimes( char* p, TimePair& tp ) {
   p = printTime( p, tp.a, false );
   *p++ = ' ';
   *p++ = '-';
   *p++ = ' ';
   return printTime( p, tp.b );
}

//----------------------------------------------------------------------------
// print the day, substituting 'Su' on Sundays
//
// d = 1..31
//
static char* printDay( char* p, long jd, int d, bool eom ) {
  if( (jd+d) % 7 == 6)        /* Sunday */
    strcpy( p, "Su" );
  else {
    int nextDay = eom ? 1 : d+1;
    sprintf( p, "%2d", nextDay );
  }
  return nextColumn(p+2,2);
}

//----------------------------------------------------------------------------
// figure out darkest hours & print to buffer
//
char* printDarkness( char* p, int i, const TimePair* const astTwi, const TimePair* const moonRSOrg )
{
  TimePair dark = { astTwi[i].TP_END, astTwi[i+1].TP_START };

  // first time in, create a local copy that we can munge
  static TimePair moonRS[DAYS];
  if ( 0 == i )
    memcpy( moonRS, moonRSOrg, sizeof(TimePair)*DAYS);

  // define day + time vars to deal with 'yesterday' and 'tomorrow'
  double darkStart = astTwi[i].TP_END + i;
  double darkEnd = astTwi[i+1].TP_START + (i+1);

  double moonRise;
  if ( moonRS[i].TP_RISE < 0. || moonRS[i].TP_SET > moonRS[i].TP_RISE) {
      moonRise = moonRS[i+1].TP_RISE + (i+1);
      moonRS[i].TP_RISE = moonRS[i+1].TP_RISE;
  }
  else
      moonRise = moonRS[i].TP_RISE + i;

  double moonSet;
  if ( moonRS[i].TP_SET < 0. || moonRS[i].TP_RISE > moonRS[i].TP_SET ) {
    moonSet = moonRS[i+1].TP_SET + (i+1);
    moonRS[i].TP_SET = moonRS[i+1].TP_SET;
  }
  else
    moonSet = moonRS[i].TP_SET + i;

  // check moon rise & set
  if (moonSet > darkStart && moonSet < darkEnd) {
    darkStart = moonSet;
    dark.TP_START = moonRS[i].TP_SET;
  }
  if (moonRise > darkStart && moonRise < darkEnd ) {
    darkEnd = moonRise;
    dark.TP_END = moonRS[i].TP_RISE;
  }

  bool noDarkness = (moonRise < darkStart && moonSet > darkEnd );

  // print out darkness range or 'none'
  if ( noDarkness ) {
    strcpy(p, " -- none -- " );
    p = nextColumn(p+12, 3);
  }
  else
    p = printTimes( p, dark );

  return p;
}

//----------------------------------------------------------------------------
// helper fn for printEvents
inline char* append(char* pTo, const char* pFrom, int len)
{
  memcpy( pTo, pFrom, len );
  return pTo + len;
}

//----------------------------------------------------------------------------
// check for lunar and solar quarters & print if found
//

char* printEvents( char*p, int i, double* jd, ObsInfo& oi, DST dstDay )
{
  char* pStart = p;

  static PlanetData pd;
  double lunarLon[2], solarLon[2];

  // get ecliptic longitude for earth & moon
  //
  for( int j=0; j<2; j++ ) {
    pd.calc( EARTH, jd[i] + double(j), oi );
    solarLon[j] = pd.eclipticLon();

    pd.calc( LUNA, jd[i] + double(j), oi );
    lunarLon[j] = pd.eclipticLon();
  }

  // We don't bother finding the exact instant of the following events.
  // The code just checks for a quadrant change and reports the  event.

  // check for lunar quarters
  //
  int quad1 = RiseSet::quadrant( lunarLon[1] - solarLon[1] );
  int quad0 = RiseSet::quadrant( lunarLon[0] - solarLon[0] );
  if( quad1 != quad0 ) {
      static const char* strings[4] = { "1Q ", "FM ", "3Q ", "NM " };
      p = append( p, strings[quad0], 3);
  }

  // check for solar quarters
  //
  quad1 = RiseSet::quadrant( solarLon[1] );
  quad0 = RiseSet::quadrant( solarLon[0] );
  if( quad1 != quad0 ) {
      static const char* strings[4] =
          { "SumSol ", "Aut Eq  ", "WinSol ", "Ver Eq " };

      p = append( p, strings[quad0], 7 );
  }

  // handle DST indicator
  if ( DST_START == dstDay )
    p = append( p, "DSTime", 6 );
  else if ( DST_END == dstDay )
    p = append( p, "STime", 5 );

  return nextColumn(p, 10-(p-pStart), (p==pStart));
}

//----------------------------------------------------------------------------
// The main program function
//
void calcAndPrint(CalData& cd)
{
  // calc. start and end days
  //
  long jd_start = DateOps::dmyToDay( 1, cd.month, cd.year );
  long jd_end = ( cd.month < 12 ) ?
      DateOps::dmyToDay( 1, cd.month + 1, cd.year ) :
      DateOps::dmyToDay( 1, 1, cd.year + 1 );

  int end = int(jd_end - jd_start);

  // fill in data for month in question
  //
  TimePair sunRS[DAYS], moonRS[DAYS], astTwi[DAYS];
  double jd[DAYS];
  static const double hourFraction = 1./24.;

  double tzAdj = (double)cd.loc.timeZone() * hourFraction;
  long dstStart = DateOps::dstStart( cd.year );
  long dstEnd = DateOps::dstEnd( cd.year );

  #if defined( PROGRESS_BAR )
    fprintf( stderr, "working" );
  #endif

  for( int i=0; i<=end+1; i++ )
  {
    long day = jd_start + i;

    // automatically adjust for DST if enabled
    // This 'rough' method will be off by one on moon rise/set between
    //   midnight and 2:00 on "clock change" days. (sun & astTwi never
    //   occur at these times.)
    //
    double dstAdj =
        ( false == g_ignoreDst && day>=dstStart && day<dstEnd) ?
        hourFraction : 0.;

    jd[i] = (double)day - (tzAdj + dstAdj) - .5;

    // calculate rise/set times for the sun
    RiseSet::getTimes( sunRS[i], RiseSet::SUN, jd[i], cd.loc );

    // calculate rise/set times for Astronomical Twilight
    RiseSet::getTimes( astTwi[i], RiseSet::ASTRONOMICAL_TWI, jd[i], cd.loc );

    // calculate rise/set time for Luna )
    RiseSet::getTimes( moonRS[i], RiseSet::MOON, jd[i], cd.loc );

    #if defined( PROGRESS_BAR )
      fputc( '.', stderr );
    #endif
  }
  fputc( '\n', stderr );

  printHeading(cd);

  // print data for each day
  //
  char buf[256];
  for( int i=0; i<end; i++ ) {

    if (g_html) {
      if ( !(i&1) )
        fprintf( g_fp, "<TR CLASS=\"bar\">\n  <TD>" );
      else
        fprintf( g_fp, "<TR>\n  <TD>" );
    }

    // print day
    char* p = printDay( buf, jd_start, i, false );

    // print darkest hours
    p = printDarkness(p, i, astTwi, moonRS);

    // check for lunar & solar quarters and DST, print if found
    DST dstDay = DST_NONE;
    if ( dstStart == jd_start+i )
      dstDay = DST_START;
    else if ( dstEnd == jd_start+i )
      dstDay = DST_END;

    p = printEvents(p, i, jd, cd.loc, dstDay);

    // print rise/set times for Luna
    p = printTime( p, moonRS[i].TP_RISE );
    p = printTime( p, moonRS[i].TP_SET );

    // print set time for the sun
    p = printTime( p, sunRS[i].TP_SET );

    // print end of Astronomical Twilight */
    p = printTime( p, astTwi[i].TP_END );

    // next day
    p = printDay( p, jd_start, i+1, i == end-1 );
    if (!g_tabDelimited && !g_html)
      *p++ = ' ';

    // print start of Astronomical Twilight */
    p = printTime( p, astTwi[i+1].TP_START );

    // print rise time for the sun     (last column)
    p = printTime( p, sunRS[i+1].TP_RISE, false );

    if (g_html)
      strcpy( p, "</TD>\n</TR>\n" );
    else {
      *p++ = '\n';
      *p = 0;
    }

    fputs( buf, g_fp );
  }
  if (g_html)
    fputs( "</TABLE>\n</BODY>\n</HTML>\n", g_fp );
}

//----------------------------------------------------------------------------
// print usage and exit
//
void usage(const char* pn, bool x = true)
{
  fprintf( stderr,
      "usage: %s <month> <year> [-d] [-h] [-t] [-u]\n"
      "       -d = ignore daylight savings time\n"
      "       -h = HTML table output\n"
      "       -t = tab-delimited output\n"
      "       -u = output time in UTC\n\n"
      "Notes:\n"
      " - Command-line options override config file settings.\n"
      " - To produce a sample config file with all options listed, run the\n"
      "   program without a %s.cfg file in the current directory.\n"
      " - This program currently only supports the Gregorian calendar.", pn, pn );
  exit(-1);
}

//----------------------------------------------------------------------------
//**** main ****
//----------------------------------------------------------------------------
int main( int argc, char* *argv )
{
  // check / get arguments
  //
  if ( argc < 3 )
    usage(argv[0]);
  if ( '-' == argv[1][0] )
    usage( argv[0] );

  CalData cd;

  char cfgFile[256];
  char etcCfgFile[256];
  char* pCfgFile = cfgFile;
  sprintf( cfgFile, "%s%s", argv[0], CFG_EXT );
  sprintf( etcCfgFile, "C:\\etc\\%s", cfgFile );

  ConfigFile cf( pCfgFile );
  if ( ConfigFile::OK != cf.status() ) {
    pCfgFile = etcCfgFile;
    // try etc
    if ( ConfigFile::OK != cf.filename( pCfgFile ) ) {
      fprintf( stderr,
          "\nWarning: unable to find %s or %s:\n\n"
          "- I'll try to create a 'template' file for you in the current directory.\n"
          "- Edit this file to reflect your location.\n ", cfgFile, etcCfgFile );
      FILE* fp = fopen( cfgFile, "w" );
      if ( NULL != fp ) {
        fprintf( fp,
            "# %s\n# Note: East and North are positive\n"
            "# e.g., Philadelphia, PA, US is latitude -75.16, longitude 39.95,\n"
            "#       and timeZone -5\n"
            "longitude=0.0\nlatitude=0.0\ntimeZone=0\n\n"
            "# Set this to true to ignore Daylight time:\n"
            "ignoreDST=false\n\n"
            "# Set this to true to use UTC (this overrides timeZone & DST)\n"
            "useUTC=false\n\n"
            "# Set this to true to output to an HTML file:\n"
            "htmlOutput=false\n\n"
            "# Set this to true to produce a text file with tab-delimited fields:\n"
            "# (Note: html & tabs are mutually exclusive!)\n"
            "tabDelimited=false\n", cfgFile );
        fclose( fp );
        exit(1);
      }
    }
  }
  printf( "Using %s.\n", pCfgFile );

  if ( cf.value( "longitude" ) )
    cd.loc.setLongitude( cf.dblValue( "longitude" ) );

  if ( cf.value( "latitude" ) )
    cd.loc.setLatitude( cf.dblValue( "latitude" ) );

  if ( cf.value( "timeZone" ) )
    cd.loc.setTimeZone( cf.intValue( "timeZone" ) );

  if ( cf.value( "htmlOutput" ) )
    g_html = cf.boolValue( "htmlOutput" );

  if ( cf.value( "ignoreDST" ) )
    g_ignoreDst = cf.boolValue( "ignoreDST" );

  if ( cf.value( "tabDelimited" ) )
    g_tabDelimited = cf.boolValue( "tabDelimited" );

  if ( cf.boolValue( "useUTC" ) ) {
    cd.loc.setTimeZone(0);
    g_ignoreDst = true;
  }

  cd.month = atoi( argv[1] );
  cd.year = atoi( argv[2] );
  if ( cd.month < 1 || cd.month > 12 || cd.year <= 0)
    usage(argv[0]);

  if ( argc > 3 ) {
    for (int i = 3; i < argc; i++ ) {
      if ( '-' != argv[i][0] )
        usage( argv[0] );
      else if ( 'd' == argv[i][1] )
        g_ignoreDst = true;
      else if ( 'h' == argv[i][1] )
        g_html = true;
      else if ( 't' == argv[i][1] )
        g_tabDelimited = true;
      else if ( 'u' == argv[i][1] ) {
        cd.loc.setTimeZone(0);
        g_ignoreDst = true;
      }
      else
        usage( argv[0] );
    }
  }

  if ( g_tabDelimited && g_html ) {
    fprintf( stderr, "Error: html and tabDelimited cannot both be specified.\n" );
    exit(-1);
  }

#if DEBUG
  if (g_tabDelimited)
    fprintf( stderr, "tab delimited\n");
  if (g_html)
    fprintf( stderr, "HTML output\n");
  if (g_ignoreDst) {
    if ( 0 == cd.loc.timeZone() )
      fprintf( stderr, "use UTC\n");
    else
      fprintf( stderr, "ignore DST\n");
  }
#endif
  if ( 0. == cd.loc.longitude() && 0. == cd.loc.latitude() )
    printf( "Latitude & Longitude are both set to 0. Is this what you intended?\n" );

  const char* pExt = 0;
  if (g_html)
    pExt = "html";
  else if (g_tabDelimited)
    pExt = "txt";

  if ( 0 != pExt ) {
    char filename[32];
    sprintf( filename, "%s%d.%s", monthNames[cd.month-1], cd.year, pExt );
    FILE* fp = fopen( filename, "w" );
    if ( NULL != fp ) {
      g_fp = fp;
      fprintf( stderr, "Writing output to %s\n", filename );
    }
  }

  calcAndPrint( cd );

  if ( stdout != g_fp )
    fclose( g_fp );

  return 0;
}
//----------------------------------------------------------------------------
