# Basic astronomical functions library - statically linked Win32 version

all:  astephem.exe persian.exe jd.exe relativi.exe tables.exe \
      get_test.exe ssattest.exe lun_test.exe lunar32.lib

lunar32.lib: \
      alt_az.obj astfuncs.obj big_vsop.obj classel.obj com_file.obj \
      cospar.obj date.obj de_plan.obj delta_t.obj dist_pa.obj elp82dat.obj \
      getplane.obj get_time.obj jsats.obj lunar2.obj triton.obj  \
      miscell.obj nutation.obj obliquit.obj pluto.obj precess.obj  \
      refract.obj refract4.obj \
      rocks.obj showelem.obj ssats.obj vislimit.obj vsopson.obj
   del lunar32.lib
   lib /OUT:lunar32.lib \
      alt_az.obj astfuncs.obj big_vsop.obj classel.obj com_file.obj \
      cospar.obj date.obj de_plan.obj delta_t.obj dist_pa.obj elp82dat.obj \
      getplane.obj get_time.obj jsats.obj lunar2.obj triton.obj  \
      miscell.obj nutation.obj obliquit.obj pluto.obj precess.obj  \
      refract.obj refract4.obj \
      rocks.obj showelem.obj ssats.obj vislimit.obj vsopson.obj >> err

jd.exe:  jd.obj lunar32.lib
   link jd.obj lunar32.lib

relativi.exe:  relativi.obj lunar32.lib
   link relativi.obj lunar32.lib

ssattest.exe:  ssattest.obj lunar32.lib
   link ssattest.obj lunar32.lib

astephem.exe:  astephem.obj eart2000.obj mpcorb.obj lunar32.lib
   link        astephem.obj eart2000.obj mpcorb.obj lunar32.lib

tables.exe: tables.obj riseset3.obj lunar32.lib
   link     tables.obj riseset3.obj lunar32.lib

lun_test.exe: lun_test.obj lun_tran.obj riseset3.obj lunar32.lib
   link       lun_test.obj lun_tran.obj riseset3.obj lunar32.lib

persian.exe: persian.obj solseqn.obj lunar32.lib
   link persian.obj solseqn.obj lunar32.lib

get_test.exe: get_test.obj lunar32.lib
   link get_test.obj lunar32.lib


.cpp.obj:
   cl /c /nologo /Ox /W3 /nologo $<

de_plan.obj: de_plan.cpp
   cl /c /nologo /W3 de_plan.cpp

relativi.obj:
   cl /c /nologo /Od /W3 /DTEST_CODE relativi.cpp

ssats.obj:
   cl /c /nologo /Od /W3 ssats.cpp

