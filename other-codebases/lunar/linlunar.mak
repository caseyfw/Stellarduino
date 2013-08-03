all: astcheck astephem calendar colors colors2 \
	cosptest dist easter get_test htc20b jd \
 jevent jpl2b32 jsattest lun_test marstime oblitest persian  \
 phases ps_1996 ssattest tables \
	test_ref testprec uranus1 utc_test

CC=g++

CFLAGS=-Wall -O3

.cpp.o:
	$(CC) $(CFLAGS) -c $<

OBJS= alt_az.o astfuncs.o big_vsop.o classel.o cospar.o date.o delta_t.o \
	de_plan.o dist_pa.o eart2000.o elp82dat.o getplane.o get_time.o \
	jsats.o lunar2.o miscell.o nutation.o obliquit.o pluto.o precess.o \
	showelem.o ssats.o triton.o vsopson.o

lunar.a: $(OBJS)
	rm -f lunar.a
	ar rv lunar.a $(OBJS)

astcheck:  astcheck.o mpcorb.o lunar.a
	$(CC) $(CFLAGS) -o astcheck astcheck.o mpcorb.o lunar.a

astephem:  astephem.o mpcorb.o lunar.a
	$(CC) $(CFLAGS) -o astephem astephem.o mpcorb.o lunar.a

calendar: calendar.o lunar.a
	$(CC) $(CFLAGS) -o calendar   calendar.o   lunar.a

colors: colors.cpp
	$(CC) $(CFLAGS) -o colors colors.cpp -DSIMPLE_TEST_PROGRAM

colors2: colors2.cpp
	$(CC) $(CFLAGS) -o colors2 colors2.cpp -DTEST_FUNC

cosptest: cosptest.o lunar.a
	$(CC) $(CFLAGS) -o cosptest   cosptest.o   lunar.a

dist: dist.cpp
	$(CC) $(CFLAGS) -o dist dist.cpp

easter: easter.cpp lunar.a
	$(CC) $(CFLAGS) -o easter -DTEST_CODE easter.cpp lunar.a

get_test: get_test.o lunar.a
	$(CC) $(CFLAGS) -o get_test get_test.o lunar.a

htc20b: htc20b.cpp lunar.a
	$(CC) $(CFLAGS) -o htc20b -DTEST_MAIN htc20b.cpp lunar.a

jd: jd.o lunar.a
	$(CC) $(CFLAGS) -o jd jd.o lunar.a

jevent:	jevent.o lunar.a
	$(CC) $(CFLAGS) -o jevent jevent.o lunar.a

jpl2b32:	jpl2b32.o
	$(CC) $(CFLAGS) -o jpl2b32 jpl2b32.o

jsattest: jsattest.o lunar.a
	$(CC) $(CFLAGS) -o jsattest jsattest.o lunar.a

lun_test:                lun_test.o lun_tran.o riseset3.o lunar.a
	$(CC)	$(CFLAGS) -o lun_test lun_test.o lun_tran.o riseset3.o lunar.a

marstime: marstime.cpp
	$(CC) $(CFLAGS) -o marstime marstime.cpp -DTEST_PROGRAM

oblitest: oblitest.o obliqui2.o spline.o lunar.a
	$(CC) $(CFLAGS) -o oblitest oblitest.o obliqui2.o spline.o lunar.a

persian: persian.o solseqn.o lunar.a
	$(CC) $(CFLAGS) -o persian persian.o solseqn.o lunar.a

phases: phases.o lunar.a
	$(CC) $(CFLAGS) -o phases   phases.o   lunar.a

ps_1996: ps_1996.o lunar.a
	$(CC) $(CFLAGS) -o ps_1996   ps_1996.o   lunar.a

relativi: relativi.cpp lunar.a
	$(CC) $(CFLAGS) -o relativi -DTEST_CODE relativi.cpp lunar.a

ssattest: ssattest.o lunar.a
	$(CC) $(CFLAGS) -o ssattest ssattest.o lunar.a

tables:                    tables.o riseset3.o lunar.a
	$(CC) $(CFLAGS) -o tables tables.o riseset3.o lunar.a

test_ref:                    test_ref.o refract.o refract4.o
	$(CC) $(CFLAGS) -o test_ref test_ref.o refract.o refract4.o

testprec:                    testprec.o lunar.a
	$(CC) $(CFLAGS) -o testprec testprec.o lunar.a

uranus1: uranus1.o gust86.o
	$(CC) $(CFLAGS) -o uranus1 uranus1.o gust86.o

utc_test:                utc_test.o lunar.a
	$(CC) $(CFLAGS) -o utc_test utc_test.o lunar.a

