import java.util.HashMap;

/**
 * Sheet: Premium (CALCULATION ENGINE)
 * Monthly premium allocation across up to ProjYears*12 months (max rows 2–1201).
 * Columns:
 *   A=Date(serial), B=Year, C=Month, D=Age,
 *   E=Gross Prem, F=Ann Prem YTD, G=Cum Prem,
 *   H=Target Rem, I=Prem Tgt, J=Prem Exc,
 *   K=Load Amt, L=Net Prem
 *
 * Row 1 = headers. Data rows start at row 2.
 *
 * The Excel EDATE formula increments the date by 1 month each row.
 * We replicate this with a simple calendar calculation based on Excel serial dates.
 *
 * B(row>=3): =IF(B_prev >= ProjYears, "", IF(C_prev=12, B_prev+1, B_prev))
 * NOTE: Row 2 uses "=1" (not conditional), Row 3 uses ">", Row 4+ uses ">=".
 * We handle both branches identically (>=) since it makes no practical difference
 * inside the active projection window.
 */
public class Sheet_Premium {

    // Max months to compute = 1200 data rows
    private static final int MAX_ROWS = 1200;

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();

    // Internal arrays for fast cross-sheet lookup (row index = excelRow - 2)
    private final double[] colA_date = new double[MAX_ROWS]; // Excel date serial
    private final Object[] colB_year = new Object[MAX_ROWS]; // Integer or ""
    private final double[] colC_month= new double[MAX_ROWS];
    private final double[] colD_age  = new double[MAX_ROWS];
    private final double[] colE_gPrem= new double[MAX_ROWS];
    private final double[] colF_ytd  = new double[MAX_ROWS];
    private final double[] colG_cum  = new double[MAX_ROWS];
    private final double[] colH_tgtRem=new double[MAX_ROWS];
    private final double[] colI_pTgt = new double[MAX_ROWS];
    private final double[] colJ_pExc = new double[MAX_ROWS];
    private final double[] colK_load = new double[MAX_ROWS];
    private final double[] colL_netP = new double[MAX_ROWS];

    private final Sheet_Inputs inputs;

    public Sheet_Premium(Sheet_Inputs inputs) {
        this.inputs = inputs;
    }

    public void initialize_calculation() {
        double issueDate  = inputs.getIssueDate();
        double issueAge   = inputs.getIssueAge();
        double projYears  = inputs.getProjYears();
        double schedPrem  = inputs.getSchedPrem();
        double targetPrem = inputs.getTargetPrem();
        double loadTgt    = inputs.getLoadTgt();
        double loadExc    = inputs.getLoadExc();

        // ── Row 2 (index 0): seed row ─────────────────────────────────────────
        colA_date[0] = issueDate;                    // A2 = Input_IssueDate
        colB_year[0] = 1.0;                          // B2 = 1
        colC_month[0] = 1.0;                         // C2 = 1
        colD_age[0]  = issueAge + (1.0 - 1.0);       // D2 = IssueAge+(B2-1)=IssueAge
        colE_gPrem[0] = schedPrem;                   // E2 = Input_SchedPrem
        colF_ytd[0]  = schedPrem;                    // F2 = E2
        colG_cum[0]  = schedPrem;                    // G2 = E2
        colH_tgtRem[0] = targetPrem;                 // H2 = Input_TargetPrem
        colI_pTgt[0] = Math.min(schedPrem, targetPrem);              // I2 = MIN(E2,H2)
        colJ_pExc[0] = schedPrem - colI_pTgt[0];                    // J2 = E2-I2
        colK_load[0] = colI_pTgt[0]*loadTgt + colJ_pExc[0]*loadExc; // K2
        colL_netP[0] = schedPrem - colK_load[0];                     // L2

        // ── Rows 3–1201 (indices 1–1199) ─────────────────────────────────────
        for (int idx = 1; idx < MAX_ROWS; idx++) {
            // A = EDATE(A_prev, 1)
            colA_date[idx] = edateSerial(colA_date[idx-1], 1);

            // B: seed row2 uses "=1", row3 uses ">", row4+ uses ">=".
            // For generality we use >= throughout (matches Excel behavior for year>=2):
            Object prevB = colB_year[idx-1];
            double prevBval = (prevB instanceof Double) ? (double) prevB : -1.0;
            double prevC = colC_month[idx-1];
            if (prevBval >= projYears) {
                colB_year[idx] = ""; // blank
            } else if (prevC == 12.0) {
                colB_year[idx] = prevBval + 1.0;
            } else {
                colB_year[idx] = prevBval;
            }

            // C: month counter
            colC_month[idx] = (prevC == 12.0) ? 1.0 : prevC + 1.0;

            // Check if active
            boolean active = !(colB_year[idx] instanceof String);
            double bVal = active ? (double) colB_year[idx] : 0.0;

            // D: age
            colD_age[idx] = issueAge + (bVal - 1.0);

            // E: Gross Prem
            colE_gPrem[idx] = active ? schedPrem : 0.0;

            // F: Ann Prem YTD
            if (!active) {
                colF_ytd[idx] = 0.0;
            } else if (colC_month[idx] == 1.0) {
                colF_ytd[idx] = colE_gPrem[idx];
            } else {
                colF_ytd[idx] = colF_ytd[idx-1] + colE_gPrem[idx];
            }

            // G: Cumulative Premium
            colG_cum[idx] = active ? (colG_cum[idx-1] + colE_gPrem[idx]) : 0.0;

            // H: Target Remaining
            if (!active) {
                colH_tgtRem[idx] = 0.0;
            } else if (colC_month[idx] == 1.0) {
                colH_tgtRem[idx] = targetPrem;
            } else {
                colH_tgtRem[idx] = Math.max(0.0, targetPrem - colF_ytd[idx-1]);
            }

            // I: Prem Tgt = MIN(E, H)
            colI_pTgt[idx] = active ? Math.min(colE_gPrem[idx], colH_tgtRem[idx]) : 0.0;

            // J: Prem Exc = E - I
            colJ_pExc[idx] = active ? (colE_gPrem[idx] - colI_pTgt[idx]) : 0.0;

            // K: Load Amount
            colK_load[idx] = active ? (colI_pTgt[idx]*loadTgt + colJ_pExc[idx]*loadExc) : 0.0;

            // L: Net Prem = E - K
            colL_netP[idx] = active ? (colE_gPrem[idx] - colK_load[idx]) : 0.0;
        }

        // ── Populate HashMaps ─────────────────────────────────────────────────
        HashMap<Integer,Object> hA=new HashMap<>(), hB=new HashMap<>(), hC=new HashMap<>(),
                hD=new HashMap<>(), hE=new HashMap<>(), hF=new HashMap<>(), hG=new HashMap<>(),
                hH=new HashMap<>(), hI=new HashMap<>(), hJ=new HashMap<>(), hK=new HashMap<>(),
                hL=new HashMap<>();
        // Headers (row 1)
        hA.put(1,"Date"); hB.put(1,"Year"); hC.put(1,"Month"); hD.put(1,"Age");
        hE.put(1,"Gross Prem"); hF.put(1,"Ann Prem YTD"); hG.put(1,"Cum Prem");
        hH.put(1,"Target Rem"); hI.put(1,"Prem Tgt"); hJ.put(1,"Prem Exc");
        hK.put(1,"Load Amt"); hL.put(1,"Net Prem");

        for (int idx = 0; idx < MAX_ROWS; idx++) {
            int row = idx + 2;
            hA.put(row, colA_date[idx]);
            hB.put(row, colB_year[idx]);
            hC.put(row, colC_month[idx]);
            hD.put(row, colD_age[idx]);
            hE.put(row, colE_gPrem[idx]);
            hF.put(row, colF_ytd[idx]);
            hG.put(row, colG_cum[idx]);
            hH.put(row, colH_tgtRem[idx]);
            hI.put(row, colI_pTgt[idx]);
            hJ.put(row, colJ_pExc[idx]);
            hK.put(row, colK_load[idx]);
            hL.put(row, colL_netP[idx]);
        }

        data.put("A", hA); data.put("B", hB); data.put("C", hC); data.put("D", hD);
        data.put("E", hE); data.put("F", hF); data.put("G", hG); data.put("H", hH);
        data.put("I", hI); data.put("J", hJ); data.put("K", hK); data.put("L", hL);
    }

    public HashMap<String, HashMap<Integer, Object>> getData() { return data; }

    // ── Fast accessors for cross-sheet use ────────────────────────────────────

    /** Get Net Prem (col L) by Excel row (2-based). */
    public double getNetPrem(int excelRow) {
        return colL_netP[excelRow - 2];
    }

    /** Get Year (col B) by Excel row. Returns "" if beyond projection. */
    public Object getYear(int excelRow) {
        return colB_year[excelRow - 2];
    }

    /** Get Month (col C) by Excel row. */
    public double getMonth(int excelRow) {
        return colC_month[excelRow - 2];
    }

    /** Get Date serial (col A) by Excel row. */
    public double getDate(int excelRow) {
        return colA_date[excelRow - 2];
    }

    /** Get Age (col D) by Excel row. */
    public double getAge(int excelRow) {
        return colD_age[excelRow - 2];
    }

    /** Get Cum Prem (col G) by Excel row. */
    public double getCumPrem(int excelRow) {
        return colG_cum[excelRow - 2];
    }

    /** Get Gross Prem (col E) by Excel row. */
    public double getGrossPrem(int excelRow) {
        return colE_gPrem[excelRow - 2];
    }

    /**
     * VLOOKUP(dateSer, Premium!A:L, 12, TRUE) — find the net prem for a date.
     * Since Premium dates are monotonically increasing, we do an exact match
     * (dates from Cash Value col A are derived from the same EDATE sequence).
     */
    public double getNetPremByDate(double dateSer) {
        // Exact match scan (dates are sequential and unique per row)
        for (int i = 0; i < MAX_ROWS; i++) {
            if (Math.abs(colA_date[i] - dateSer) < 0.5) {
                return colL_netP[i];
            }
        }
        // Fallback: approximate match — return last row where date <= dateSer
        double result = 0;
        for (int i = 0; i < MAX_ROWS; i++) {
            if (colA_date[i] <= dateSer + 0.5) result = colL_netP[i];
            else break;
        }
        return result;
    }

    // ── EDATE helper ─────────────────────────────────────────────────────────

    /**
     * Adds 'months' months to an Excel date serial.
     * Replicates Excel EDATE: preserves day-of-month clamped to month end.
     */
    public static double edateSerial(double serial, int months) {
        // Convert Excel serial → year/month/day
        int[] ymd = serialToYMD(serial);
        int year = ymd[0], month = ymd[1], day = ymd[2];
        month += months;
        while (month > 12) { month -= 12; year++; }
        while (month <  1) { month += 12; year--; }
        int maxDay = daysInMonth(year, month);
        if (day > maxDay) day = maxDay;
        return ymdToSerial(year, month, day);
    }

    /** Convert Excel date serial to {year, month, day}. */
    private static int[] serialToYMD(double serial) {
        // Excel serial: 1 = 1900-01-01, with the Lotus 1900 leap-year bug
        // (serial 60 is treated as 1900-02-29, which doesn't exist)
        int s = (int) serial;
        if (s >= 60) s--; // adjust for Lotus bug
        // Days since 1899-12-31
        int year = 1899, month = 12, day = 31;
        // Use a simple proleptic algorithm
        // Julian Day Number approach
        long jd = s + 2415018L; // offset to Julian Day
        long z = jd + 1;
        long a = (long)((z - 1867216.25) / 36524.25);
        long aa = z + 1 + a - a/4;
        long b = aa + 1524;
        long c = (long)((b - 122.1) / 365.25);
        long d = (long)(365.25 * c);
        long e = (long)((b - d) / 30.6001);
        day   = (int)(b - d - (long)(30.6001 * e));
        month = (int)(e < 14 ? e - 1 : e - 13);
        year  = (int)(month > 2 ? c - 4716 : c - 4715);
        return new int[]{year, month, day};
    }

    /** Convert {year, month, day} to Excel date serial. */
    private static double ymdToSerial(int year, int month, int day) {
        // Reverse of serialToYMD
        long jd;
        if (month <= 2) { year--; month += 12; }
        long a = year / 100;
        long b = 2 - a + a/4;
        jd = (long)(365.25*(year + 4716)) + (long)(30.6001*(month+1)) + day + b - 1524;
        long s = jd - 2415018L;
        // Undo Lotus bug: serials >= 60 need +1
        if (s >= 60) s++;
        return (double) s;
    }

    private static int daysInMonth(int year, int month) {
        if (month == 2) {
            return ((year%4==0 && year%100!=0) || year%400==0) ? 29 : 28;
        }
        return new int[]{0,31,28,31,30,31,30,31,31,30,31,30,31}[month];
    }
}
