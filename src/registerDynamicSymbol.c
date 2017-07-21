#include <R.h>
#include <Rinternals.h>
#include <stdlib.h> // for NULL
#include <R_ext/Rdynload.h>

/* FIXME: 
   Check these declarations against the C/Fortran source code.
*/

/* .Call calls */
extern SEXP _officer_a_border(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_a_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_a_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_css_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_css_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_p_ph(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_pml_run_pic(SEXP, SEXP, SEXP);
extern SEXP _officer_pml_table(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_rpr_css(SEXP);
extern SEXP _officer_rpr_new(SEXP);
extern SEXP _officer_rpr_p(SEXP);
extern SEXP _officer_rpr_w(SEXP);
extern SEXP _officer_w_ppr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_w_tcpr(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);
extern SEXP _officer_wml_run_pic(SEXP, SEXP, SEXP);
extern SEXP _officer_wml_table(SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP, SEXP);

static const R_CallMethodDef CallEntries[] = {
    {"_officer_a_border",    (DL_FUNC) &_officer_a_border,     6},
    {"_officer_a_ppr",       (DL_FUNC) &_officer_a_ppr,       15},
    {"_officer_a_tcpr",      (DL_FUNC) &_officer_a_tcpr,      19},
    {"_officer_css_ppr",     (DL_FUNC) &_officer_css_ppr,     15},
    {"_officer_css_tcpr",    (DL_FUNC) &_officer_css_tcpr,    19},
    {"_officer_p_ph",        (DL_FUNC) &_officer_p_ph,         9},
    {"_officer_pml_run_pic", (DL_FUNC) &_officer_pml_run_pic,  3},
    {"_officer_pml_table",   (DL_FUNC) &_officer_pml_table,    8},
    {"_officer_rpr_css",     (DL_FUNC) &_officer_rpr_css,      1},
    {"_officer_rpr_new",     (DL_FUNC) &_officer_rpr_new,      1},
    {"_officer_rpr_p",       (DL_FUNC) &_officer_rpr_p,        1},
    {"_officer_rpr_w",       (DL_FUNC) &_officer_rpr_w,        1},
    {"_officer_w_ppr",       (DL_FUNC) &_officer_w_ppr,       15},
    {"_officer_w_tcpr",      (DL_FUNC) &_officer_w_tcpr,      19},
    {"_officer_wml_run_pic", (DL_FUNC) &_officer_wml_run_pic,  3},
    {"_officer_wml_table",   (DL_FUNC) &_officer_wml_table,    8},
    {NULL, NULL, 0}
};

void R_init_officer(DllInfo *dll)
{
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
