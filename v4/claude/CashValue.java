package claude;

/**
 * Cash Value sheet  (rows 2-121, i.e. 120 monthly periods)
 *
 * Columns (Excel letters → field names):
 *   A  date        B  year        C  month       D  age
 *   ─── NONG ──────────────────────────────────────────────
 *   F  begAV       G  netPrem     H  charges      I  interest
 *   J  endAV       K  surrChg     L  csv          M  db
 *   ─── GUAR ──────────────────────────────────────────────
 *   O  gBegAV      P  gNetPrem    Q  gCharges     R  gInterest
 *   S  gEndAV
 *
 * Key design notes
 * ─────────────────
 * 1.  BegAV  (F)  seed row: = ACV (from Inputs)
 *     GUAR BegAV (O) is ALSO = ACV in every row (the formula is literally
 *     "=ACV" in all rows of col O, not chained from the prior row's gEndAV).
 *
 * 2.  VLOOKUP(D, Table_Rates, colIndex, FALSE) → Rates.lookupXxx(age)
 *       col 2 → COI_Rate      → Rates.lookupCOI(age)
 *       col 3 → SurrChg_Rate  → Rates.lookupSurrChg(age)
 *       col 4 → Corridor_Factor → Rates.lookupCorridor(age)
 *
 * 3.  VLOOKUP(A, Premium!rows, 12, TRUE) picks up the net premium (col L of
 *     Premium) for the matching date.  Because dates advance by 1 month each
 *     row and are identical between the two sheets, we simply pass the Premium
 *     array in and index by row number.
 *
 * 4.  GUAR charges use COI * 1.5 (guaranteed mortality loading).
 */
public class CashValue {

    public static final int ROWS = 120;

    // ── NONG ────────────────────────────────────────────────────────────────
    public final double[] year;
    public final boolean[] yearBlank;
    public final double[] month;
    public final double[] age;
    public final double[] begAV;
    public final double[] netPrem;
    public final double[] charges;
    public final double[] interest;
    public final double[] endAV;
    public final double[] surrChg;
    public final double[] csv;
    public final double[] db;

    // ── GUAR ────────────────────────────────────────────────────────────────
    public final double[] gBegAV;
    public final double[] gNetPrem;
    public final double[] gCharges;
    public final double[] gInterest;
    public final double[] gEndAV;

    public CashValue(Premium prem) {
        year      = new double[ROWS];
        yearBlank = new boolean[ROWS];
        month     = new double[ROWS];
        age       = new double[ROWS];
        begAV     = new double[ROWS];
        netPrem   = new double[ROWS];
        charges   = new double[ROWS];
        interest  = new double[ROWS];
        endAV     = new double[ROWS];
        surrChg   = new double[ROWS];
        csv       = new double[ROWS];
        db        = new double[ROWS];
        gBegAV    = new double[ROWS];
        gNetPrem  = new double[ROWS];
        gCharges  = new double[ROWS];
        gInterest = new double[ROWS];
        gEndAV    = new double[ROWS];

        compute(prem);
    }

    private void compute(Premium prem) {

        // ── Row 0 (Excel row 2) – seed row ───────────────────────────────────
        // year/month/age mirror the Premium sheet seed
        year[0]      = 1.0;
        yearBlank[0] = false;
        month[0]     = 1.0;
        age[0]       = Inputs.Input_IssueAge + (year[0] - 1.0);

        // F2: =ACV
        begAV[0] = Inputs.ACV;

        // G2: VLOOKUP(A2, Premium!1:1048, 12, TRUE)  → prem.netPrem[0]
        netPrem[0] = prem.netPrem[0];

        // H2: =(COI(age) * MAX(0, FaceAmt - (BegAV+NetPrem))) + AdminFee
        charges[0] = computeChargeNONG(age[0], begAV[0], netPrem[0]);

        // I2: =(BegAV+NetPrem-Charges) * MthlyRate_Cur
        interest[0] = (begAV[0] + netPrem[0] - charges[0]) * Inputs.Calc_MthlyRate_Cur;

        // J2: =MAX(0, BegAV+NetPrem-Charges+Interest)
        endAV[0] = Math.max(0, begAV[0] + netPrem[0] - charges[0] + interest[0]);

        // K2: =SurrChg(age) * FaceAmt
        surrChg[0] = Rates.lookupSurrChg(age[0]) * Inputs.Input_FaceAmt;

        // L2: =MAX(0, EndAV - SurrChg)
        csv[0] = Math.max(0, endAV[0] - surrChg[0]);

        // M2: =MAX(FaceAmt, EndAV * Corridor(age))
        db[0] = Math.max(Inputs.Input_FaceAmt, endAV[0] * Rates.lookupCorridor(age[0]));

        // ─── GUAR seed ────────────────────────────────────────────────────────
        // O2: =ACV  (literal ACV in every row, NOT chained from prior gEndAV)
        gBegAV[0] = Inputs.ACV;

        // P2: =Premium!L2
        gNetPrem[0] = prem.netPrem[0];

        // Q2: =(COI(age)*1.5 * MAX(0, FaceAmt-(gBegAV+gNetPrem))) + AdminFee
        gCharges[0] = computeChargeGUAR(age[0], gBegAV[0], gNetPrem[0]);

        // R2: =(gBegAV+gNetPrem-gCharges) * MthlyRate_Guar
        gInterest[0] = (gBegAV[0] + gNetPrem[0] - gCharges[0]) * Inputs.Calc_MthlyRate_Guar;

        // S2: =MAX(0, gBegAV+gNetPrem-gCharges+gInterest)
        gEndAV[0] = Math.max(0, gBegAV[0] + gNetPrem[0] - gCharges[0] + gInterest[0]);

        // ── Rows 1+ (Excel rows 3+) ───────────────────────────────────────────
        for (int i = 1; i < ROWS; i++) {
            double prevYear  = year[i - 1];
            double prevMonth = month[i - 1];
            boolean prevBlank = yearBlank[i - 1];

            // B (year) – same logic as Premium sheet
            if (i == 1) {
                // row 3: IF(B2 > ProjYears, "", ...)
                if (prevYear > Inputs.Input_ProjYears) {
                    yearBlank[i] = true; year[i] = 0;
                } else {
                    yearBlank[i] = false;
                    year[i] = (prevMonth == 12) ? prevYear + 1 : prevYear;
                }
            } else {
                // row 4+: IF(B_prev >= ProjYears, "", ...)
                if (prevBlank || prevYear >= Inputs.Input_ProjYears) {
                    yearBlank[i] = true; year[i] = 0;
                } else {
                    yearBlank[i] = false;
                    year[i] = (prevMonth == 12) ? prevYear + 1 : prevYear;
                }
            }

            month[i] = (prevMonth == 12) ? 1 : prevMonth + 1;
            age[i]   = yearBlank[i] ? 0 : Inputs.Input_IssueAge + (year[i] - 1.0);

            if (yearBlank[i]) {
                // All columns = 0 when year is blank
                begAV[i] = netPrem[i] = charges[i] = interest[i] = 0;
                endAV[i] = surrChg[i] = csv[i] = db[i] = 0;
                gBegAV[i] = gNetPrem[i] = gCharges[i] = gInterest[i] = gEndAV[i] = 0;
                continue;
            }

            // F: =IF(B="", 0, J_prev)   (prior row's NONG EndAV)
            begAV[i] = endAV[i - 1];

            // G: VLOOKUP(date, Premium, 12, TRUE) → prem.netPrem[i]
            netPrem[i] = prem.netPrem[i];

            // H: charges (NONG)
            charges[i] = computeChargeNONG(age[i], begAV[i], netPrem[i]);

            // I: interest (NONG)
            interest[i] = (begAV[i] + netPrem[i] - charges[i]) * Inputs.Calc_MthlyRate_Cur;

            // J: endAV (NONG)
            endAV[i] = Math.max(0, begAV[i] + netPrem[i] - charges[i] + interest[i]);

            // K: surrender charge
            surrChg[i] = Rates.lookupSurrChg(age[i]) * Inputs.Input_FaceAmt;

            // L: CSV
            csv[i] = Math.max(0, endAV[i] - surrChg[i]);

            // M: DB
            db[i] = Math.max(Inputs.Input_FaceAmt, endAV[i] * Rates.lookupCorridor(age[i]));

            // ─── GUAR ─────────────────────────────────────────────────────────
            // O: =ACV  (always the same constant in every row per the Excel formula)
            gBegAV[i] = Inputs.ACV;

            // P: =Premium!Lrow
            gNetPrem[i] = prem.netPrem[i];

            // Q: charges (GUAR)
            gCharges[i] = computeChargeGUAR(age[i], gBegAV[i], gNetPrem[i]);

            // R: interest (GUAR)
            gInterest[i] = (gBegAV[i] + gNetPrem[i] - gCharges[i]) * Inputs.Calc_MthlyRate_Guar;

            // S: gEndAV
            gEndAV[i] = Math.max(0, gBegAV[i] + gNetPrem[i] - gCharges[i] + gInterest[i]);
        }
    }

    // ── Charge helpers ────────────────────────────────────────────────────────

    /**
     * NONG charges (H column):
     *   = (COI(age) * MAX(0, FaceAmt - (BegAV + NetPrem))) + AdminFee
     */
    private static double computeChargeNONG(double age, double bav, double np) {
        double netAmount = Math.max(0, Inputs.Input_FaceAmt - (bav + np));
        return Rates.lookupCOI(age) * netAmount + Inputs.Input_AdminFee;
    }

    /**
     * GUAR charges (Q column):
     *   = ((COI(age) * 1.5) * MAX(0, FaceAmt - (gBegAV + gNetPrem))) + AdminFee
     */
    private static double computeChargeGUAR(double age, double gbav, double gnp) {
        double netAmount = Math.max(0, Inputs.Input_FaceAmt - (gbav + gnp));
        return (Rates.lookupCOI(age) * 1.5) * netAmount + Inputs.Input_AdminFee;
    }
}