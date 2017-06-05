#include <R.h>
#include <Rinternals.h>
#include <stdlib.h> // for NULL
#include <R_ext/Rdynload.h>

/* FIXME:
Check these declarations against the C/Fortran source code.
*/

/* .Call calls */
extern SEXP officer_a_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_a_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_css_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_css_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_p_ph(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_pml_run_pic(SEXP, SEXP, SEXP);
extern SEXP officer_pml_table(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_rpr_css(SEXP);
extern SEXP officer_rpr_new(SEXP);
extern SEXP officer_rpr_p(SEXP);
extern SEXP officer_rpr_w(SEXP);
extern SEXP officer_w_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_w_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP officer_wml_run_pic(SEXP, SEXP, SEXP);
extern SEXP officer_wml_table(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);

static const R_CallMethodDef CallEntries[] = {
  {"officer_a_ppr",       (DL_FUNC) &officer_a_ppr,       15},
  {"officer_a_tcpr",      (DL_FUNC) &officer_a_tcpr,      19},
  {"officer_css_ppr",     (DL_FUNC) &officer_css_ppr,     15},
  {"officer_css_tcpr",    (DL_FUNC) &officer_css_tcpr,    19},
  {"officer_p_ph",        (DL_FUNC) &officer_p_ph,         9},
  {"officer_pml_run_pic", (DL_FUNC) &officer_pml_run_pic,  3},
  {"officer_pml_table",   (DL_FUNC) &officer_pml_table,    8},
  {"officer_rpr_css",     (DL_FUNC) &officer_rpr_css,      1},
  {"officer_rpr_new",     (DL_FUNC) &officer_rpr_new,      1},
  {"officer_rpr_p",       (DL_FUNC) &officer_rpr_p,        1},
  {"officer_rpr_w",       (DL_FUNC) &officer_rpr_w,        1},
  {"officer_w_ppr",       (DL_FUNC) &officer_w_ppr,       15},
  {"officer_w_tcpr",      (DL_FUNC) &officer_w_tcpr,      19},
  {"officer_wml_run_pic", (DL_FUNC) &officer_wml_run_pic,  3},
  {"officer_wml_table",   (DL_FUNC) &officer_wml_table,    8},
  {NULL, NULL, 0}
};

void R_init_officer(DllInfo *dll)
{
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
}
