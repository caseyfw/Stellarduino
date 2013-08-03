#define MODEL_N_PARAMS     17
#define MODEL_ALT_RATE     0
#define MODEL_ALT_ZERO     1
#define MODEL_AZ_RATE      2
#define MODEL_AZ_ZERO      3
#define MODEL_TILT1        4
#define MODEL_TILT2        5
#define MODEL_MISALIGN     6
#define MODEL_FLEXURE      7
#define MODEL_AZ_COS       8
#define MODEL_AZ_SIN       9
#define MODEL_ALT_SIN      10
#define MODEL_ALT_COS2     11
#define MODEL_ALT_SIN2     12
#define MODEL_AZ_COS2      13
#define MODEL_AZ_SIN2      14
#define MODEL_WARPED_TBL_1 15
#define MODEL_WARPED_TBL_2 16


int encoder_to_alt_az( const double *m, const double x, const double y,
                           double *alt, double *az);
int alt_az_to_encoder( const double *m, const double alt, const double az,
                           double *x, double *y);
int set_up_scope_model( double *m, const  int n_stars, const double *x,
               const double *y, const double *alt, const double *az);
int improve_model( double *m, const int n_stars, const double *x,
            const double *y, const double *alt, const double *az,
            const char *adjustable);
