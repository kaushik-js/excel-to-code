import java.util.HashMap;

/**
 * Sheet_Output - OUTPUT sheet
 *
 * Translates the "Output" sheet (A1:H19) from gaulEngine_v2.xlsx.
 *
 * Columns:
 *   A  Policy Year       (1, 2, 3, ... up to ProjYears, then "")
 *   B  End Date          (FILTER: Premium date where Month=12 AND Year=A)
 *   C  Attained Age      (VLOOKUP B2 → Premium col 4 = Age)
 *   D  Annual Premium    (SUM(FILTER: Premium GrossPrem where Year=A))
 *   E  Cumulative Premium (VLOOKUP B → Premium col 7 = CumPrem)
 *   F  NONG Cash Value   (FILTER: CashValue EndAV where Date=B)
 *   G  NONG Death Benefit (FILTER: CashValue DB where Date=B)
 *   H  GUAR Cash Value   (FILTER: CashValue GuarEndAV where Date=B)
 *
 * Rows 2–19 (18 data rows max).
 *
 * Array-formula FILTER emulation:
 *   - B: scan Premium A2:A121 to find date where C=12 AND B=yearVal
 *   - D: sum Premium E2:E121 where B=yearVal
 *   - E: VLOOKUP(B, Premium A:L, 7) → Premium CumPrem for date B
 *   - F/G/H: scan CashValue A2:A121 to find date matching B
 */
public class Sheet_Output {

    private final Sheet_Inputs    inputs;
    private final Sheet_Premium   premium;
    private final Sheet_CashValue cashValue;

    // Internal arrays (0-indexed → Excel row = index + 2), 18 output rows
    private static final int ROWS = 18;

    private final Object[] colA_year;   // Double or ""
    private final Object[] colB_endDate; // Double serial or ""
    private final Object[] colC_age;    // Double or ""
    private final Object[] colD_annPrem; // Double or ""
    private final Object[] colE_cumPrem; // Double or ""
    private final Object[] colF_nongCV; // Double or ""
    private final Object[] colG_nongDB; // Double or ""
    private final Object[] colH_guarCV; // Double or ""

    public Sheet_Output(Sheet_Inputs inputs, Sheet_Premium premium, Sheet_CashValue cashValue) {
        this.inputs    = inputs;
        this.premium   = premium;
        this.cashValue = cashValue;

        colA_year    = new Object[ROWS];
        colB_endDate = new Object[ROWS];
        colC_age     = new Object[ROWS];
        colD_annPrem = new Object[ROWS];
        colE_cumPrem = new Object[ROWS];
        colF_nongCV  = new Object[ROWS];
        colG_nongDB  = new Object[ROWS];
        colH_guarCV  = new Object[ROWS];
    }

    public void initialize_calculation() {
        int projYears = (int) inputs.getProjYears();

        for (int i = 0; i < ROWS; i++) {
            int excelRow = i + 2;

            // ── Column A: Policy Year ────────────────────────────────────────
            // A2: =1
            // A3: =IF(A2 < Input_ProjYears-1, A2+1, "")   (rows 3-12: < ProjYrs-1)
            // A13+: various slightly different boundary comparisons in Excel
            // Simplify: output year i+1 if i+1 <= ProjYears, else ""
            if (i == 0) {
                colA_year[i] = 1.0;
            } else {
                Object prevYear = colA_year[i - 1];
                if (prevYear instanceof String) {
                    colA_year[i] = "";
                } else {
                    double py = (Double) prevYear;
                    // The Excel formulas use slightly inconsistent boundary checks
                    // (rows 3-11 use <ProjYears-1, rows 12-14 use <=ProjYears-1,
                    //  rows 15-19 use <=ProjYears). The effect is: output years 1..ProjYears
                    // The simplest faithful translation: year = py+1 if py+1 <= projYears
                    colA_year[i] = (py + 1 <= projYears) ? (py + 1) : "";
                }
            }

            if (colA_year[i] instanceof String) {
                // Beyond projection — all cells are blank
                colB_endDate[i] = "";
                colC_age[i]     = "";
                colD_annPrem[i] = "";
                colE_cumPrem[i] = "";
                colF_nongCV[i]  = "";
                colG_nongDB[i]  = "";
                colH_guarCV[i]  = "";
                continue;
            }

            double yearVal = (Double) colA_year[i];

            // ── Column B: End Date ───────────────────────────────────────────
            // FILTER(Premium!A$2:A$121, (Premium!C$2:C$121=12)*(Premium!B$2:B$121=yearVal))
            // → find first Premium row (rows 2-121) where Month=12 AND Year=yearVal
            double endDate = 0;
            int endPremRow = -1;
            for (int r = 2; r <= 121; r++) {
                Object yr = premium.getYear(r);
                if (!(yr instanceof Double)) continue;
                if ((Double) yr == yearVal && premium.getMonth(r) == 12.0) {
                    endDate = premium.getDate(r);
                    endPremRow = r;
                    break;
                }
            }
            colB_endDate[i] = (endPremRow >= 0) ? endDate : "";

            // ── Column C: Attained Age ───────────────────────────────────────
            // VLOOKUP(B, Premium!A2:L121, 4) → Age column
            if (endPremRow >= 0) {
                colC_age[i] = premium.getAge(endPremRow);
            } else {
                colC_age[i] = "";
            }

            // ── Column D: Annual Premium ─────────────────────────────────────
            // SUM(FILTER(Premium!E$2:E$121, Premium!B$2:B$121=yearVal))
            double sumPrem = 0;
            for (int r = 2; r <= 121; r++) {
                Object yr = premium.getYear(r);
                if (yr instanceof Double && (Double) yr == yearVal) {
                    sumPrem += premium.getGrossPrem(r);
                }
            }
            colD_annPrem[i] = sumPrem;

            // ── Column E: Cumulative Premium ─────────────────────────────────
            // VLOOKUP(B, Premium!A2:L121, 7) → CumPrem (col G = index 7)
            if (endPremRow >= 0) {
                colE_cumPrem[i] = premium.getCumPrem(endPremRow);
            } else {
                colE_cumPrem[i] = "";
            }

            // ── Columns F/G/H: CashValue lookups by matching date = B ────────
            // FILTER(CashValue!J$2:J$121, CashValue!A$2:A$121 = B)
            if (endPremRow >= 0) {
                int cvRow = cashValue.findRowByDate(endDate);
                if (cvRow >= 0 && cvRow <= 121) {
                    colF_nongCV[i] = cashValue.getNongEndAV(cvRow);
                    colG_nongDB[i] = cashValue.getNongDB(cvRow);
                    colH_guarCV[i] = cashValue.getGuarEndAV(cvRow);
                } else {
                    colF_nongCV[i] = "";
                    colG_nongDB[i] = "";
                    colH_guarCV[i] = "";
                }
            } else {
                colF_nongCV[i] = "";
                colG_nongDB[i] = "";
                colH_guarCV[i] = "";
            }
        }
    }

    // ── Full HashMap output ───────────────────────────────────────────────────
    public HashMap<String, HashMap<Integer, Object>> getOutput() {
        HashMap<String, HashMap<Integer, Object>> sheet = new HashMap<>();

        HashMap<Integer, Object> colA = new HashMap<>();
        HashMap<Integer, Object> colB = new HashMap<>();
        HashMap<Integer, Object> colC = new HashMap<>();
        HashMap<Integer, Object> colD = new HashMap<>();
        HashMap<Integer, Object> colE = new HashMap<>();
        HashMap<Integer, Object> colF = new HashMap<>();
        HashMap<Integer, Object> colG = new HashMap<>();
        HashMap<Integer, Object> colH = new HashMap<>();

        // Row 1: headers
        colA.put(1, "Policy Year");
        colB.put(1, "End Date");
        colC.put(1, "Attained Age");
        colD.put(1, "Annual Premium");
        colE.put(1, "Cumulative Premium");
        colF.put(1, "NONG Cash Value");
        colG.put(1, "NONG Death Benefit");
        colH.put(1, "GUAR Cash Value");

        for (int i = 0; i < ROWS; i++) {
            int row = i + 2;
            colA.put(row, colA_year[i]);
            colB.put(row, colB_endDate[i]);
            colC.put(row, colC_age[i]);
            colD.put(row, colD_annPrem[i]);
            colE.put(row, colE_cumPrem[i]);
            colF.put(row, colF_nongCV[i]);
            colG.put(row, colG_nongDB[i]);
            colH.put(row, colH_guarCV[i]);
        }

        sheet.put("A", colA);
        sheet.put("B", colB);
        sheet.put("C", colC);
        sheet.put("D", colD);
        sheet.put("E", colE);
        sheet.put("F", colF);
        sheet.put("G", colG);
        sheet.put("H", colH);

        return sheet;
    }
}
