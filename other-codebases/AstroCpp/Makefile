# Makefile for Astro Lib & DarkCal Program
# (this is for gnu make, which assumes gcc)
# remember that TAB chars (not spaces) are required before the cmd lines!

LIBOBJS = AstroOps.o DateOps.o DateOpsIntl.o PlanetData.o Lunar.o Pluto.o MathOps.o RiseSet.o Vsop.o 
#VsopData.o 
LIB = Astro.a
DC_BASE = DarkCal

# Lose the ".exe" for linux
DC = $(DC_BASE).exe

#DEBUG = -g
CPPFLAGS = $(DEBUG) -ffast-math
# The fast-math option allows GCC to violate some ANSI or IEEE rules and/or
# specifications in the interest of optimizing code for speed.  For example,
# it allows the compiler to assume arguments to the sqrt() function are
# non-negative numbers and that no floating-point values are NaNs.

# Bad things will happen if this assumption proves false...

all: $(LIB) $(DC)

clean:
	rm -f $(LIB) $(DC) $(wildcard *.o)

$(LIBOBJS): Makefile

$(LIB) : $(LIBOBJS)
	ar -r $(LIB) $(LIBOBJS)

$(DC): $(DC_BASE).o ConfigFile.o $(LIB) Makefile
	$(CXX) -o $(DC) $(DC_BASE).o ConfigFile.o $(LIB)

AstroOps.o: AstroOps.h

ConfigFile.o: ConfigFile.h

$(DC_BASE).o: AstroOps.h PlanetData.h DateOps.h RiseSet.h ConfigFile.h

DateOps.o: DateOps.h

DateOpsIntl.o: DateOps.h

PlanetData.o: PlanetData.h AstroOps.h MathOps.h Lunar.h Pluto.h Vsop.h

Lunar.o: Lunar.h LunarTerms.h

Pluto.o: Pluto.h PlutoTerms.h

MathOps.o: MathOps.h AstroOps.h

RiseSet.o: RiseSet.h AstroOps.h PlanetData.h Vsop.h

Vsop.o: Vsop.h VsopData.cpp


