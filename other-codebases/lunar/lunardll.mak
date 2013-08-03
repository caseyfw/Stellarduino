# Basic astronomical functions library - Win32 .DLL version

all:  astcheck.exe astephem.exe calendar.exe cosptest.exe dist.exe \
      easter.exe get_test.exe htc20b.exe jd.exe jevent.exe \
      jpl2b32.exe jsattest.exe lun_test.exe marstime.exe oblitest.exe \
      persian.exe phases.exe ps_1996.exe relativi.exe ssattest.exe tables.exe \
      testprec.exe test_ref.exe uranus1.exe utc_test.exe


lunar.lib: \
      alt_az.obj astfuncs.obj big_vsop.obj classel.obj com_file.obj \
      cospar.obj date.obj de_plan.obj delta_t.obj dist_pa.obj elp82dat.obj \
      getplane.obj get_time.obj jsats.obj lunar2.obj  \
      miscell.obj nutation.obj obliquit.obj pluto.obj precess.obj  \
      refract.obj refract4.obj rocks.obj showelem.obj \
      ssats.obj triton.obj vislimit.obj vsopson.obj
   del lunar.lib
   del lunar.dll
   link /DLL /MAP /IMPLIB:lunar.lib /DEF:lunar.def \
      alt_az.obj astfuncs.obj big_vsop.obj classel.obj com_file.obj \
      cospar.obj date.obj de_plan.obj delta_t.obj dist_pa.obj elp82dat.obj \
      getplane.obj get_time.obj jsats.obj lunar2.obj  \
      miscell.obj nutation.obj obliquit.obj pluto.obj precess.obj  \
      refract.obj refract4.obj rocks.obj showelem.obj \
      ssats.obj triton.obj vislimit.obj vsopson.obj >> err

astcheck.exe:  astcheck.obj eart2000.obj mpcorb.obj lunar.lib
   link astcheck.obj eart2000.obj mpcorb.obj lunar.lib

astephem.exe:  astephem.obj eart2000.obj mpcorb.obj lunar.lib
   link astephem.obj eart2000.obj mpcorb.obj lunar.lib

calendar.exe:  calendar.obj lunar.lib
   link calendar.obj lunar.lib

cosptest.exe:  cosptest.obj lunar.lib
   link cosptest.obj lunar.lib

dist.exe:  dist.obj
   link    dist.obj

easter.exe: easter.cpp
   cl -W4 -Ox -DTEST_CODE -nologo easter.cpp

get_test.exe: get_test.obj lunar.lib
   link get_test.obj lunar.lib

htc20b.exe: htc20b.cpp
   cl -W3 -Ox -DTEST_MAIN -nologo htc20b.cpp

jevent.exe: jevent.obj lunar.lib
   link     jevent.obj lunar.lib

jd.exe:  jd.obj lunar.lib
   link jd.obj lunar.lib

jpl2b32.exe:  jpl2b32.obj
   link       jpl2b32.obj

jsattest.exe: jsattest.obj lunar.lib
   link       jsattest.obj lunar.lib

lun_test.exe: lun_test.obj lun_tran.obj riseset3.obj lunar.lib
   link lun_test.obj lun_tran.obj riseset3.obj lunar.lib

marstime.exe: marstime.cpp
   cl /Ox /W3 /DTEST_PROGRAM marstime.cpp

oblitest.exe:  oblitest.obj obliqui2.obj spline.obj lunar.lib
   link        oblitest.obj obliqui2.obj spline.obj lunar.lib

persian.exe: persian.obj solseqn.obj lunar.lib
   link      persian.obj solseqn.obj lunar.lib

phases.exe:  phases.obj lunar.lib
   link      phases.obj lunar.lib

ps_1996.exe:  ps_1996.obj lunar.lib
   link       ps_1996.obj lunar.lib

relativi.exe:  relativi.obj lunar.lib
   link        relativi.obj lunar.lib

relativi.obj:
   cl /c /Od /W3 /DTEST_CODE -nologo relativi.cpp

ssattest.exe:  ssattest.obj lunar.lib
   link ssattest.obj lunar.lib

ssats.obj: ssats.cpp
   cl -W3 -Od -GX -c -LD -nologo -I\myincl ssats.cpp >> err
   type err

tables.exe: tables.obj riseset3.obj lunar.lib
   link     tables.obj riseset3.obj lunar.lib

testprec.exe:  testprec.obj lunar.lib
   link        testprec.obj lunar.lib

test_ref.exe:  test_ref.obj refract.obj refract4.obj
   link        test_ref.obj refract.obj refract4.obj

uranus1.exe:  uranus1.obj gust86.obj
   link       uranus1.obj gust86.obj

utc_test.exe:  utc_test.obj lunar.lib
   link        utc_test.obj lunar.lib

CFLAGS=-W3 -Ox -GX -c -LD -DNDEBUG -nologo

.cpp.obj:
   cl $(CFLAGS) $< >> err
   type err

