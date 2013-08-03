//----------------------------------------------------------------------------
// VisLimit - calculates the visual limiting magnitude
//
// Astro library based on open-source code from project pluto
//
// mark huss 11/2000
//----------------------------------------------------------------------------

struct FixedBrightnessData {
  // constants for a given time:
  double zenithAngMoon, zenithAngSun, moonElongation;
  double htAboveSeaInMeters, latitude;
  double temperatureInC, relativeHumidity;
  double year, month;
};

struct AngularBrightnessData {
  // values varying across the sky:
  double zenithAngle;
  double distMoon, distSun;         // angular,  not real,linear
};


class VisLimit {
public:
  int setBrightnessParams(FixedBrightnessData& fbd);
  int computeSkyBrightness(AngularBrightnessData& abd);
  double computeLimitingMag();
  int computeExtinction();

  void setMask(int m) { mask = m; }

  double getK(unsigned i) { return ( i < BANDS ) ? k[i] : -1.; }
  double getBrightness(unsigned i) { return ( i < BANDS ) ? brightness[i] : -1.; }
  double getExtinction(unsigned i) { return ( i < BANDS ) ? extinction[i] : -1.; }

private:
  enum {
    BANDS = 5
  };
  // constants for a given time:
  FixedBrightnessData fixed;

  // values varying across the sky:
  AngularBrightnessData angular;

  int mask;   // indicates which of the 5 photometric bands we want

  // Items computed in setBrightnessParams:
  double airMassSun, airMassMoon, lunarMag;
  double k[BANDS], c3[BANDS], c4[BANDS], ka[BANDS], kr[BANDS], ko[BANDS], kw[BANDS];
  double yearTerm;

  // Items computed in computeLimitingMag:
  double airMassGas, airMassAerosol, airMassOzone;
  double extinction[BANDS];

  // Internal parameters from computeSkyBrightness:
  double air_mass, brightness[5];

  static double computeAirMass( const double zenithAngle);
  static double computeFFactor( double obj_dist);

};
