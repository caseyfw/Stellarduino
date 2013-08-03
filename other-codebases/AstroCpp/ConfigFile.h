/*--------------------------------------------------------------------------*\
 * ConfigFile class
 *
 * Input file consists of lines of text of the form
 *   name = value
 * where 'name' is the look up key and 'value' is the resultant value
 *
 * '\' is the escape character and allows values to contain embedded '"' and
 *   '#' characters, as well as to do line continuation.
 * To put '"' in a value use '\"'.  To put '\' in a value use '\\'.
 *
 * Anything in double quotes is copied literally (except escape sequences).
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

#if !defined(CONFIG_FILE__H)
#define CONFIG_FILE__H

#include <stdio.h>  // for FILE*

//----------------------------------------------------------------------------
// CFList : a simple linked list class
//
class CFList {
public:
  CFList(const char* name, const char* value): m_next(0) {
    m_name = strDup(name);
    m_value = strDup(value);
  }

  ~CFList() {
    delete [] m_name; m_name=0;
    delete [] m_value; m_value=0;

    // recursive delete
    delete m_next;
    m_next = 0;
  }

  void add(const char* name, const char* value);
  void update(const char* name, const char* value);

  const char* value(const char* name) const;
  bool boolValue(const char* name) const;
  int intValue(const char* name) const;
  double dblValue(const char* name) const;

  void dump( FILE* fp = stdout ) const;

private:
  CFList() {}  // default c'tor not allowed

  static char* strDup( const char* p );

  const char* m_name;
  const char* m_value;
  CFList* m_next;
};  // end class CFList

//----------------------------------------------------------------------------
// the main class

class ConfigFile {

public:
  // cheap way to define integer constants
  enum {
    NV_SEP_CHAR = '=',
    COMMENT_CHAR = '#',
    QUOTE_CHAR = '"',
    ESC_CHAR = '\\',
    BUFSIZE = 5120
  };

  enum {
    OK = 0,
    INVALID_FILE_NAME,
    FILE_NOT_FOUND,
    NO_ENTRIES
  };

  static bool debug;

  // constructors
  ConfigFile() : m_status( NO_ENTRIES ), m_items( "ConfigFile", "init" ) {}

  ConfigFile( const char* fname ) : m_items( "ConfigFile", "init" ) {
     readAndParseFile( fname );
  }

  // specify a new config file
  int filename( const char* fname ) { return readAndParseFile( fname ); }

  // retrieve a string value --return NULL if not found
  const char* value( const char* name ) const  {
    return m_items.value( name );
  }

  // retrieve a boolean value -- returns false if not found or name is not '[Tt]rue'
  int boolValue( const char* name ) const {
    return m_items.boolValue( name );
  }

  // retrieve an int value -- returns 0 if not found or name is not an int
  int intValue( const char* name ) const {
    return m_items.intValue( name );
  }

  // retrieve a double value -- returns 0. if not found or name is not a double
  double dblValue( const char* name ) const {
    return m_items.dblValue( name );
  }

  // output all keys and values
  void dump( FILE* fp ) const { m_items.dump( fp ); }

  int status() const { return m_status; }

private:
  int m_status;
  CFList m_items;

  // config file path & name
  const char* m_filename;

  // load m_values
  int readAndParseFile( const char* fname );

  bool compactBuffer(char* buf);

}; // end class ConfigFile

#endif  /* #if !defined(CONFIG_FILE__H) */
