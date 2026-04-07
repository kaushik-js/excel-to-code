import java.util.HashMap;

/**
 * Sheet_CashValue - CALCULATION ENGINE
 *
 * Translates the "Cash Value" sheet (A1:S1201) from gaulEngine_v2.xlsx.
 *
 * Columns:
 *   A  Date         (Excel serial, mirrors Premium!A)
 *   B  Year         (mirrors Premium!B, "" when beyond projection)
 *   C  Month        (mirrors Premium!C)
 *   D  Age          (mirrors Premium!D)
 *   E  --- NONG --- (separator, always empty)
 *   F  Beg AV       (NONG beginning account value)
 *   G  Net Prem     (NONG = VLOOKUP date → Premium col L)
 *   H  Charges      (NONG = COI * MAX(0, FaceAmt - (BegAV+NetPrem)) + AdminFee)
 *   I  Interest     (NONG = (BegAV+NetPrem-Charges) * MthlyRate_Cur)
 *   J  End AV       (NONG = MAX(0, BegAV+NetPrem-Charges+Interest))
 *   K  Surr Chg     (NONG = SurrChgRate * FaceAmt)
 *   L  CSV          (NONG = MAX(0, EndAV - SurrChg))
 *   M  DB           (NONG = MAX(FaceAmt, EndAV * CorridorFactor))
 *   N  --- GUAR --- (separator)
 *   O  Beg AV       (GUAR, always = ACV for rows 2+)
 *   P  Net Prem     (GUAR = Premium!L for that row)
 *   Q  Charges      (GUAR = COI*1.5 * MAX(0, FaceAmt-(BegAV+NetPrem)) + AdminFee)
 *   R  Interest     (GUAR = (BegAV+NetPrem-Charges) * MthlyRate_Guar)
 *   S  End AV       (GUAR = MAX(0, BegAV+NetPrem-Charges+Interest))
 *
 * Note: GUAR Beg AV is always ACV (the formula =ACV in every row is literal).
 * Rows 2–1201 (Excel), i.e. 1200 monthly rows.
 */
public class Sheet_CashValue {

    private final Sheet_Inputs  inputs;
    private final Sheet_Rates   rates;
    private final Sheet_Premium premium;

    // Internal arrays (0-indexed → Excel row = index + 2)
    private final double[] colA_date;   // Date serial
    private final Object[] colB_year;   // Year (Double or "")
    private final double[] colC_month;  // Month
    private final double[] colD_age;    // Age

    // NONG
    private final double[] colF_begAV;
    private final double[] colG_netPrem;
    private final double[] colH_charges;
    private final double[] colI_interest;
    private final double[] colJ_endAV;
    private final double[] colK_surrChg;
    private final double[] colL_csv;
    private final double[] colM_db;

    // GUAR
    private final double[] colO_begAV;
    private final double[] colP_netPrem;
    private final double[] colQ_charges;
    private final double[] colR_interest;
    private final double[] colS_endAV;

    private static final int ROWS = 1200;

    public Sheet_CashValue(Sheet_Inputs inputs, Sheet_Rates rates, Sheet_Premium premium) {
        this.inputs  = inputs;
        this.rates   = rates;
        this.premium = premium;

        colA_date    = new double[ROWS];
        colB_year    = new Object[ROWS];
        colC_month   = new double[ROWS];
        colD_age     = new double[ROWS];

        colF_begAV   = new double[ROWS];
        colG_netPrem = new double[ROWS];
        colH_charges = new double[ROWS];
        colI_interest = new double[ROWS];
        colJ_endAV   = new double[ROWS];
        colK_surrChg = new double[ROWS];
        colL_csv     = new double[ROWS];
        colM_db      = new double[ROWS];

        colO_begAV   = new double[ROWS];
        colP_netPrem = new double[ROWS];
        colQ_charges = new double[ROWS];
        colR_interest = new double[ROWS];
        colS_endAV   = new double[ROWS];
    }

    public void initialize_calculation() {
        double faceAmt      = inputs.getFaceAmt();
        double acv          = inputs.getACV();
        double adminFee     = inputs.getAdminFee();
        double mthlyRateCur = inputs.getMthlyRateCur();
        double mthlyRateGuar = inputs.getMthlyRateGuar();
        int    projYears    = (int) inputs.getProjYears();

        // ── Row 2 (index 0) ──────────────────────────────────────────────────
        // A2: =Input_IssueDate
        colA_date[0] = inputs.getIssueDate();
        // B2: =1
        colB_year[0] = 1.0;
        // C2: =1
        colC_month[0] = 1.0;
        // D2: =Input_IssueAge+(B2-1) = IssueAge + 0
        colD_age[0] = inputs.getIssueAge();

        // F2: =ACV
        colF_begAV[0] = acv;
        // G2: =VLOOKUP(A2, Premium!1:1048, 12, TRUE) → net prem for this date
        colG_netPrem[0] = premium.getNetPremByDate(colA_date[0]);
        // H2: charges
        colH_charges[0] = calcNongCharges(colD_age[0], colF_begAV[0], colG_netPrem[0],
                                           faceAmt, adminFee);
        // I2: interest
        colI_interest[0] = (colF_begAV[0] + colG_netPrem[0] - colH_charges[0]) * mthlyRateCur;
        // J2: end AV
        colJ_endAV[0] = Math.max(0, colF_begAV[0] + colG_netPrem[0] - colH_charges[0] + colI_interest[0]);
        // K2: surr chg
        colK_surrChg[0] = rates.getSurrChgRate((int) colD_age[0]) * faceAmt;
        // L2: CSV
        colL_csv[0] = Math.max(0, colJ_endAV[0] - colK_surrChg[0]);
        // M2: death benefit
        colM_db[0] = Math.max(faceAmt, colJ_endAV[0] * rates.getCorridorFactor((int) colD_age[0]));

        // O2: =ACV
        colO_begAV[0] = acv;
        // P2: =Premium!L2
        colP_netPrem[0] = premium.getNetPrem(2);
        // Q2: charges (guar uses COI*1.5)
        colQ_charges[0] = calcGuarCharges(colD_age[0], colO_begAV[0], colP_netPrem[0],
                                           faceAmt, adminFee);
        // R2: interest
        colR_interest[0] = (colO_begAV[0] + colP_netPrem[0] - colQ_charges[0]) * mthlyRateGuar;
        // S2: end AV
        colS_endAV[0] = Math.max(0, colO_begAV[0] + colP_netPrem[0] - colQ_charges[0] + colR_interest[0]);

        // ── Rows 3–1201 (indices 1–1199) ────────────────────────────────────
        for (int i = 1; i < ROWS; i++) {
            int excelRow = i + 2; // Excel row for this index

            // A: EDATE(prev, 1)
            colA_date[i] = Sheet_Premium.edateSerial(colA_date[i - 1], 1);

            // B: year progression
            // Row 3: =IF(B2>Input_ProjYears,"", IF(C2=12, B2+1, B2))
            // Row 4+: =IF(B_prev>=Input_ProjYears,"", IF(C_prev=12, B_prev+1, B_prev))
            Object prevYear = colB_year[i - 1];
            double prevMonth = colC_month[i - 1];
            if (prevYear instanceof String) {
                colB_year[i] = "";
            } else {
                double prevYearD = (Double) prevYear;
                // Row 3 uses >, rows 4+ use >=  (matches Excel formulas exactly)
                boolean cutoff = (i == 1) ? (prevYearD > projYears) : (prevYearD >= projYears);
                if (cutoff) {
                    colB_year[i] = "";
                } else {
                    colB_year[i] = (prevMonth == 12) ? prevYearD + 1 : prevYearD;
                }
            }

            // C: month cycle
            colC_month[i] = (prevMonth == 12) ? 1.0 : prevMonth + 1.0;

            // D: age = IssueAge + (Year - 1)
            if (colB_year[i] instanceof String) {
                colD_age[i] = 0;
            } else {
                colD_age[i] = inputs.getIssueAge() + ((Double) colB_year[i] - 1);
            }

            boolean beyondProj = (colB_year[i] instanceof String);

            // ── NONG ──
            // F: BegAV = EndAV of previous row (or 0 if beyond proj)
            colF_begAV[i] = beyondProj ? 0 : colJ_endAV[i - 1];
            // G: net prem via date lookup in Premium
            colG_netPrem[i] = beyondProj ? 0 : premium.getNetPremByDate(colA_date[i]);
            // H: charges
            colH_charges[i] = beyondProj ? 0 :
                calcNongCharges(colD_age[i], colF_begAV[i], colG_netPrem[i], faceAmt, adminFee);
            // I: interest
            colI_interest[i] = beyondProj ? 0 :
                (colF_begAV[i] + colG_netPrem[i] - colH_charges[i]) * mthlyRateCur;
            // J: end AV
            colJ_endAV[i] = beyondProj ? 0 :
                Math.max(0, colF_begAV[i] + colG_netPrem[i] - colH_charges[i] + colI_interest[i]);
            // K: surr chg
            colK_surrChg[i] = beyondProj ? 0 :
                rates.getSurrChgRate((int) colD_age[i]) * faceAmt;
            // L: CSV
            colL_csv[i] = beyondProj ? 0 : Math.max(0, colJ_endAV[i] - colK_surrChg[i]);
            // M: death benefit
            colM_db[i] = beyondProj ? 0 :
                Math.max(faceAmt, colJ_endAV[i] * rates.getCorridorFactor((int) colD_age[i]));

            // ── GUAR ──
            // O: Beg AV is always =ACV (literal in every row per Excel formula)
            colO_begAV[i] = acv;
            // P: Net Prem from Premium sheet (same Excel row)
            colP_netPrem[i] = beyondProj ? 0 : premium.getNetPrem(excelRow);
            // Q: charges
            colQ_charges[i] = beyondProj ? 0 :
                calcGuarCharges(colD_age[i], colO_begAV[i], colP_netPrem[i], faceAmt, adminFee);
            // R: interest
            colR_interest[i] = beyondProj ? 0 :
                (colO_begAV[i] + colP_netPrem[i] - colQ_charges[i]) * mthlyRateGuar;
            // S: end AV
            colS_endAV[i] = beyondProj ? 0 :
                Math.max(0, colO_begAV[i] + colP_netPrem[i] - colQ_charges[i] + colR_interest[i]);
        }
    }

    // ── Helper: NONG charges ──────────────────────────────────────────────────
    // H = (COI_Rate * MAX(0, FaceAmt - (BegAV+NetPrem))) + AdminFee
    private double calcNongCharges(double age, double begAV, double netPrem,
                                    double faceAmt, double adminFee) {
        double coiRate = rates.getCOIRate((int) age);
        return (coiRate * Math.max(0, faceAmt - (begAV + netPrem))) + adminFee;
    }

    // ── Helper: GUAR charges ─────────────────────────────────────────────────
    // Q = (COI_Rate * 1.5 * MAX(0, FaceAmt - (BegAV+NetPrem))) + AdminFee
    private double calcGuarCharges(double age, double begAV, double netPrem,
                                    double faceAmt, double adminFee) {
        double coiRate = rates.getCOIRate((int) age);
        return ((coiRate * 1.5) * Math.max(0, faceAmt - (begAV + netPrem))) + adminFee;
    }

    // ── Public accessors for Sheet_Output ────────────────────────────────────

    /** Get NONG End AV (col J) by Excel row. */
    public double getNongEndAV(int excelRow) {
        return colJ_endAV[excelRow - 2];
    }

    /** Get NONG Death Benefit (col M) by Excel row. */
    public double getNongDB(int excelRow) {
        return colM_db[excelRow - 2];
    }

    /** Get GUAR End AV (col S) by Excel row. */
    public double getGuarEndAV(int excelRow) {
        return colS_endAV[excelRow - 2];
    }

    /** Get Date serial (col A) by Excel row. */
    public double getDate(int excelRow) {
        return colA_date[excelRow - 2];
    }

    /**
     * Find the first Cash Value row index (0-based) whose date matches the given serial.
     * Used by Output sheet to look up F/G/H by matching date from Premium.
     */
    public int findRowByDate(double dateSer) {
        for (int i = 0; i < ROWS; i++) {
            if (Math.abs(colA_date[i] - dateSer) < 0.5) return i + 2; // return Excel row
        }
        return -1;
    }

    // ── Full HashMap output ───────────────────────────────────────────────────
    public HashMap<String, HashMap<Integer, Object>> getOutput() {
        HashMap<String, HashMap<Integer, Object>> sheet = new HashMap<>();

        HashMap<Integer, Object> colA = new HashMap<>();
        HashMap<Integer, Object> colB = new HashMap<>();
        HashMap<Integer, Object> colC = new HashMap<>();
        HashMap<Integer, Object> colD = new HashMap<>();
        HashMap<Integer, Object> colF = new HashMap<>();
        HashMap<Integer, Object> colG = new HashMap<>();
        HashMap<Integer, Object> colH = new HashMap<>();
        HashMap<Integer, Object> colI = new HashMap<>();
        HashMap<Integer, Object> colJ = new HashMap<>();
        HashMap<Integer, Object> colK = new HashMap<>();
        HashMap<Integer, Object> colL = new HashMap<>();
        HashMap<Integer, Object> colM = new HashMap<>();
        HashMap<Integer, Object> colO = new HashMap<>();
        HashMap<Integer, Object> colP = new HashMap<>();
        HashMap<Integer, Object> colQ = new HashMap<>();
        HashMap<Integer, Object> colR = new HashMap<>();
        HashMap<Integer, Object> colS = new HashMap<>();

        // Row 1 headers
        colA.put(1, "Date");
        colB.put(1, "Year");
        colC.put(1, "Month");
        colD.put(1, "Age");
        colF.put(1, "Beg AV");
        colG.put(1, "Net Prem");
        colH.put(1, "Charges");
        colI.put(1, "Interest");
        colJ.put(1, "End AV");
        colK.put(1, "Surr Chg");
        colL.put(1, "CSV");
        colM.put(1, "DB");
        colO.put(1, "Beg AV");
        colP.put(1, "Net Prem");
        colQ.put(1, "Charges");
        colR.put(1, "Interest");
        colS.put(1, "End AV");

        for (int i = 0; i < ROWS; i++) {
            int row = i + 2;
            colA.put(row, colA_date[i]);
            colB.put(row, colB_year[i]);
            colC.put(row, colC_month[i]);
            colD.put(row, colD_age[i]);
            colF.put(row, colF_begAV[i]);
            colG.put(row, colG_netPrem[i]);
            colH.put(row, colH_charges[i]);
            colI.put(row, colI_interest[i]);
            colJ.put(row, colJ_endAV[i]);
            colK.put(row, colK_surrChg[i]);
            colL.put(row, colL_csv[i]);
            colM.put(row, colM_db[i]);
            colO.put(row, colO_begAV[i]);
            colP.put(row, colP_netPrem[i]);
            colQ.put(row, colQ_charges[i]);
            colR.put(row, colR_interest[i]);
            colS.put(row, colS_endAV[i]);
        }

        sheet.put("A", colA);
        sheet.put("B", colB);
        sheet.put("C", colC);
        sheet.put("D", colD);
        sheet.put("F", colF);
        sheet.put("G", colG);
        sheet.put("H", colH);
        sheet.put("I", colI);
        sheet.put("J", colJ);
        sheet.put("K", colK);
        sheet.put("L", colL);
        sheet.put("M", colM);
        sheet.put("O", colO);
        sheet.put("P", colP);
        sheet.put("Q", colQ);
        sheet.put("R", colR);
        sheet.put("S", colS);

        return sheet;
    }
}
