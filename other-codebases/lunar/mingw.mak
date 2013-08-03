all: astcheck.exe astephem.exe calendar.exe colors.exe colors2.exe \
 cosptest.exe dist.exe \
	easter.exe get_test.exe htc20b.exe jd.exe jevent.exe jpl2b32.exe \
 jsattest.exe lun_test.exe marstime.exe oblitest.exe persian.exe  \
 phases.exe ps_1996.exe relativi.exe ssattest.exe tables.exe    \
	test_ref.exe testprec.exe uranus1.exe utc_test.exe

CC=g++

CFLAGS=-Wall -O3

.cpp.o:
	$(CC) $(CFLAGS) -c $<

OBJS= alt_az.o astfuncs.o big_vsop.o classel.o com_file.o cospar.o \
	date.o delta_t.o de_plan.o dist_pa.o eart2000.o elp82dat.o getplane.o \
	get_time.o jsats.o lunar2.o miscell.o nutation.o obliquit.o pluto.o \
	precess.o showelem.o ssats.o triton.o vsopson.o

lunar.a: $(OBJS)
	del lunar.a
	ar rv lunar.a $(OBJS)

astcheck.exe:  astcheck.o mpcorb.o lunar.a
	$(CC) $(CFLAGS) -o astcheck astcheck.o mpcorb.o lunar.a

astephem.exe:  astephem.o mpcorb.o lunar.a
	$(CC) $(CFLAGS) -o astephem astephem.o mpcorb.o lunar.a

calendar.exe: calendar.o lunar.a
	$(CC) $(CFLAGS) -o calendar   calendar.o   lunar.a

colors.exe: colors.cpp
	$(CC) $(CFLAGS) -o colors colors.cpp -DSIMPLE_TEST_PROGRAM

colors2.exe: colors2.cpp
	$(CC) $(CFLAGS) -o colors2 colors2.cpp -DTEST_FUNC

cosptest.exe: cosptest.o lunar.a
	$(CC) $(CFLAGS) -o cosptest   cosptest.o   lunar.a

dist.exe: dist.cpp
	$(CC) $(CFLAGS) -o dist dist.cpp

easter.exe: easter.cpp lunar.a
	$(CC) $(CFLAGS) -o easter -DTEST_CODE easter.cpp lunar.a

get_test.exe: get_test.o lunar.a
	$(CC) $(CFLAGS) -o get_test get_test.o lunar.a

htc20b.exe: htc20b.cpp lunar.a
	$(CC) $(CFLAGS) -o htc20b -DTEST_MAIN htc20b.cpp lunar.a

jd.exe: jd.o lunar.a
	$(CC) $(CFLAGS) -o jd jd.o lunar.a

jevent.exe:	jevent.o lunar.a
	$(CC) $(CFLAGS) -o jevent jevent.o lunar.a

jpl2b32.exe:	jpl2b32.o
	$(CC) $(CFLAGS) -o jpl2b32 jpl2b32.o

jsattest.exe: jsattest.o lunar.a
	$(CC) $(CFLAGS) -o jsattest jsattest.o lunar.a

lun_test.exe:                lun_test.o lun_tran.o riseset3.o lunar.a
	$(CC)	$(CFLAGS) -o lun_test lun_test.o lun_tran.o riseset3.o lunar.a

marstime.exe: marstime.cpp
	$(CC) $(CFLAGS) -o marstime marstime.cpp -DTEST_PROGRAM

oblitest.exe: oblitest.o obliqui2.o spline.o lunar.a
	$(CC) $(CFLAGS) -o oblitest oblitest.o obliqui2.o spline.o lunar.a

persian.exe: persian.o solseqn.o lunar.a
	$(CC) $(CFLAGS) -o persian persian.o solseqn.o lunar.a

phases.exe: phases.o lunar.a
	$(CC) $(CFLAGS) -o phases   phases.o   lunar.a

ps_1996.exe: ps_1996.o lunar.a
	$(CC) $(CFLAGS) -o ps_1996   ps_1996.o   lunar.a

relativi.exe: relativi.cpp lunar.a
	$(CC) $(CFLAGS) -o relativi -DTEST_CODE relativi.cpp lunar.a

ssattest.exe: ssattest.o lunar.a
	$(CC) $(CFLAGS) -o ssattest ssattest.o lunar.a

tables.exe: tables.o riseset3.o lunar.a
	$(CC) $(CFLAGS) -o tables tables.o riseset3.o lunar.a

test_ref.exe:                test_ref.o refract.o refract4.o
	$(CC) $(CFLAGS) -o test_ref test_ref.o refract.o refract4.o

testprec.exe:                testprec.o lunar.a
	$(CC) $(CFLAGS) -o testprec testprec.o lunar.a

uranus1.exe: uranus1.o gust86.o
	$(CC) $(CFLAGS) -o uranus1 uranus1.o gust86.o

utc_test.exe:                utc_test.o lunar.a
	$(CC) $(CFLAGS) -o utc_test utc_test.o lunar.a

