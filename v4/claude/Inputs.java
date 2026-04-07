package claude;

/**
 * Inputs sheet
 * Holds all hard-coded input values from the Inputs tab.
 * Named ranges resolved directly here.
 */
public class Inputs {

    // ── Policy Setup ─────────────────────────────────────────────────────────
    public static final double Input_FaceAmt    = 100000.0;
    public static final int    Input_IssueAge   = 35;
    /**
     * Issue date as a Java "Excel serial day" (days since 1900-01-00, same as
     * Excel's date serial for 2009-03-27).
     * 2009-03-27 → Excel serial 39899.
     * We store it as a double to keep the same type as Excel date cells.
     */
    public static final double Input_IssueDate  = 39899.0; // 2009-03-27
    public static final String Input_Gender     = "Female";
    public static final String Input_Class      = "Non-Smoker";

    // ── Projection Settings ───────────────────────────────────────────────────
    public static final int    Input_ProjYears  = 11;

    // ── Premiums ──────────────────────────────────────────────────────────────
    public static final double Input_SchedPrem  = 50.0;
    public static final double Input_TargetPrem = 630.71;
    public static final double Input_GL_Annual  = 1049.05;
    public static final double Input_GL_Single  = 11801.53;

    // ── Loads & Charges ───────────────────────────────────────────────────────
    public static final double Input_Load_Tgt   = 0.07;
    public static final double Input_Load_Exc   = 0.05;

    // ── Accumulated Cash Value (ACV) ──────────────────────────────────────────
    public static final double ACV              = 11410.54;

    // ── Monthly Admin Fee ─────────────────────────────────────────────────────
    public static final double Input_AdminFee   = 5.0;

    // ── Interest Rates ────────────────────────────────────────────────────────
    public static final double Input_Int_Cur    = 0.0352;
    public static final double Input_Int_Guar   = 0.03;

    // ── Derived (named-range formulas) ────────────────────────────────────────
    /** = (1 + Input_Int_Cur)^(1/12) - 1 */
    public static final double Calc_MthlyRate_Cur  =
            Math.pow(1.0 + Input_Int_Cur,  1.0 / 12.0) - 1.0;

    /** = (1 + Input_Int_Guar)^(1/12) - 1 */
    public static final double Calc_MthlyRate_Guar =
            Math.pow(1.0 + Input_Int_Guar, 1.0 / 12.0) - 1.0;
}