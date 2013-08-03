/*--------------------------------------------------------------------------*\
 * ConfigFile class
 *
 * Input file consists of lines of text of the form
 *   name = value
 * where 'name' is the look up key and 'value' is the resultant value
 *
 * '\' is the escape character and allows values to contain embedded '"' and
 *   '#' characters.
 * To put '"' in a value use '\"'.  To put '\' in a value use '\\'.
 *
 * '\' is also the line continuation char.  If '\' is the last character on
 *   a line, the following line will be appended to that line before the
 *   line is processed.
 *
 * Anything in double quotes is copied literally, except '\', as described
 *   above.
 *
 * Outside of double quotes:
 *   '#' starts a comment; the rest of the line is ignored
 *   White space is ignored
 *
 *--------------------------------------------------------------------------
 *
 * Created by Mark Huss <mark@mhuss.com>
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
\*--------------------------------------------------------------------------*/

#include "ConfigFile.h"

using namespace std;

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

//----------------------------------------------------------------------------
// CFList methods
//----------------------------------------------------------------------------

// C++ version of strdup()
//
char* CFList::strDup( const char* s ) {
  char* p = 0;
  if (s) if (*s) {
    p = new char[strlen(s)+1];
    if (p)
      strcpy( p, s );
  }
  return p;
}

// add a name/value pair to the list
//
void CFList::add(const char* name, const char* value) {
  if (name) if (*name) if (value) if (*value) {
    CFList* p = this;
    while ( 0 != p->m_next)
      p=p->m_next;

    p->m_next = new CFList( name, value );
  }
}

// update the value of a name/value pair in the list
// if name is not found, does an add()
//
void CFList::update(const char* name, const char* value) {
  if (name) if (*name) {
    CFList* p = this;
    bool found=false;

    while (0 != p && !found) {
      if ( !strcmp(p->m_name, name) ) {
        found = true;
        delete [] p->m_value;
        p->m_value = strDup(value);
      }
      p = p->m_next;
    }
    if ( !found )
      add( name, value );
  }
}

// get the string value that corresponds the the spec'd name
// returns null string if not found
// ALWAYS RETURNS A VALID POINTER
//
const char* CFList::value(const char* name) const {
  const CFList* p = this;
  const char* pVal = "";
  if (name) if (*name) {
    while (0 != p) {
      if ( !strcmp(p->m_name, name) ) {
        if ( p->m_value )
          pVal = p->m_value;
        break;
      }
      p = p->m_next;
    }
  }
  return pVal;
}

// get the boolean value that corresponds the the spec'd name
// returns false if not found
//
bool CFList::boolValue(const char* name) const {
  const char* p = value( name );
  return ( 't' == *p || 'T' == *p || ('1' == *p && 0 == *(p+1)) );
}

// get the int value that corresponds the the spec'd name
// returns 0 if not found
//
int CFList::intValue(const char* name) const {
  const char* p = value( name );
  return (*p) ? atoi( p ) : 0;
}

// get the double value that corresponds the the spec'd name
// returns 0.0 if not found
//
double CFList::dblValue(const char* name) const {
  const char* p = value( name );
  return (*p) ? atof( p ) : 0.;
}

// output all keys and values
//
void CFList::dump( FILE* fp ) const {
  const CFList* p = this;
  while ( 0 != p ) {
    fprintf( fp, "%s = %s\n", p->m_name, p->m_value );
    p=p->m_next;
  }
}

//----------------------------------------------------------------------------
// ConfigFile methods
//----------------------------------------------------------------------------

bool ConfigFile::debug = false;

//----------------------------------------------------------------------------
// compact input buffer (one line of the input file)
// eat white space, deal with quoted strings & escaped characters
//
bool ConfigFile::compactBuffer(char* buffer) {
  char* from = buffer;
  while ( (*from) > 0 && (*from) <= ' ' ) {
    from++;
  }
  // skip 'whole line' comments and blank lines
  if ( COMMENT_CHAR == (*from) || 0 == (*from) )
    return false;

  char* to = buffer;
  bool inQuotes = false;

  while ( *from ) {
    char c = *from++;
    if ( ESC_CHAR == c ) {           // deal with escaped chars
      if ( *from )
        *to++ = *from++;
    }
    else if ( c < ' ' )             // skip \n, etc.
      continue;
    else if ( QUOTE_CHAR == c ) {   // toggle quote flag
      inQuotes = !inQuotes;
      *to++ = c;
    }
    else {
      if ( inQuotes )               // if in quotes, copy everything
        *to++ = c;
      else {                        // outside of quotes,
        if ( c <= ' ' )               // if white space, skip it
          ;
        else if ( COMMENT_CHAR == c ) // if comment, ignore rest of line
          break;
        else                        // none of the above, just copy it
          *to++ = c;
      }
    }
  }
  *to = 0;

  if( debug )
    fprintf( stderr, "    compacted line: '%s'\n", buffer);

  return true;
}

//----------------------------------------------------------------------------
// load names & values from the input file
//
int ConfigFile::readAndParseFile( const char* fname ) {
  if( debug ) {
    fprintf( stderr, "> readAndParseFile( %s )\n", fname );
  }

  m_status = NO_ENTRIES;

  m_filename = fname;

  // check for filename
  //
  if ( !fname ) {
    m_status = INVALID_FILE_NAME;
    return m_status;
  }

  // open & read file
  //
  FILE* fp = fopen( m_filename, "r" );
  if ( !fp ) {
    m_status = FILE_NOT_FOUND;
    return m_status;
  }

  char buffer[BUFSIZE];

  while ( true ) {
    memset( buffer, 0, BUFSIZE );

    if ( 0 == fgets( buffer, BUFSIZE, fp ) )
      break;

    // check for and handle line continuation chars
    int len = strlen( buffer );
    while ( true ) {
      char* pEol = &buffer[len-1];
      if ( ESC_CHAR == *pEol ) {
        if ( 0 == fgets( buffer, BUFSIZE, fp ) ) {
          *pEol = 0;
          break;
        }
      }
      else
        break;
    }
    // lose EOL
    if ( buffer[len] <= ' ' )
      buffer[len] = 0;

    if( debug )
      fprintf( stderr, "    original line: '%s'\n", buffer );

    // compact input buffer
    // eat white space, deal with quoted strings & escaped characters
    //
    if (false ==compactBuffer(buffer))
      continue;  // false means empty buffer

    // parse name/value pair and add to Hashtable
    //
    char* line = buffer;

    // minimum valid length is 3 ( 'a=b' )
    if ( strlen(line) >= 3 ) {

      // look for separator
      char* sepIndex = strchr( line, NV_SEP_CHAR );
      if (0 != sepIndex) {
        // separator char found, split into name/value
        int nameLen = sepIndex-line;
        char name[BUFSIZE];
        char value[BUFSIZE];
        memcpy( name, line, nameLen );
        name[nameLen] = 0;
        strcpy( value, sepIndex+1 );

        // remove quote chars from quoted values
        int vlen = strlen(value);
        if ( vlen > 2 &&   // minimum valid length is 3 ( '"c"' )
            QUOTE_CHAR == value[0] &&     // starts with '"'
            QUOTE_CHAR == value[vlen-1] ) // ends with '"'
        {
          strncpy(value, value+1, vlen-2 );  // cut off both ends
          value[vlen-2] = 0;
        }

        if( debug ) {
          fprintf(stderr, "    name='%s', value='%s'\n", name, value);
        }

        // add name/value pair

        m_items.add( name, value );
      }
      // ignore lines with no separators for now
    }
  }
  m_items.update( "ConfigFile", fname );
  m_status = OK;

  return m_status;
}

#if defined(UNIT_TEST)
//----------------------------------------------------------------------------
// unit test bed
//
int main( int argc, char** argv ) {

  ConfigFile cf;
  //cf.debug = true;
  int rc = 0;

  if ( argc > 1 ) {
    cf.filename( argv[1] );
    cf.dump( stdout );
  }
  else {
    fprintf( stderr, "usage: %s <config file>\n", argv[0] );
    rc = 1;
  }

  return rc;

} // end function main()

//----------------------------------------------------------------------------

#endif
