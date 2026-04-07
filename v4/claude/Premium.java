package claude;

import java.time.LocalDate;

/**
 * Premium sheet  (rows 2-121, i.e. 120 monthly periods)
 *
 * Columns (1-indexed matching Excel):
 *   A = date          (Excel serial date, not stored; derived from row index)
 *   B = year          (policy year, 1-based; empty string → 0 sentinel here)
 *   C = month         (1-12)
 *   D = age           (attained, fractional by year offset)
 *   E = grossPrem     (Gross Premium)
 *   F = annPremYTD    (Ann Prem YTD)
 *   G = cumPrem       (Cumulative Premium)
 *   H = targetRem     (Target Remaining)
 *   I = premTgt       (Premium portion up to target)
 *   J = premExc       (Premium excess of target)
 *   K = loadAmt       (Load Amount)
 *   L = netPrem       (Net Premium)
 *
 * A "blank year" in Excel (row beyond projection period) is represented here
 * as yearBlank = true; all monetary columns become 0.0 in those rows.
 */
public class Premium {

    public static final int ROWS = 120; // rows 2-121 → 120 monthly periods

    // ── Output arrays (index 0 = row 2 = period 1) ───────────────────────────
    public final double[] year;       // 0 when blank
    public final boolean[] yearBlank; // true when the Excel cell would be ""
    public final double[] month;
    public final double[] age;
    public final double[] grossPrem;
    public final double[] annPremYTD;
    public final double[] cumPrem;
    public final double[] targetRem;
    public final double[] premTgt;
    public final double[] premExc;
    public final double[] loadAmt;
    public final double[] netPrem;

    // ── Date array (Excel serial day numbers) ─────────────────────────────────
    public final double[] date;

    public Premium() {
        year       = new double[ROWS];
        yearBlank  = new boolean[ROWS];
        month      = new double[ROWS];
        age        = new double[ROWS];
        grossPrem  = new double[ROWS];
        annPremYTD = new double[ROWS];
        cumPrem    = new double[ROWS];
        targetRem  = new double[ROWS];
        premTgt    = new double[ROWS];
        premExc    = new double[ROWS];
        loadAmt    = new double[ROWS];
        netPrem    = new double[ROWS];
        date       = new double[ROWS];

        compute();
    }

    // ── EDATE helper ─────────────────────────────────────────────────────────
    /**
     * Replicates Excel EDATE(serialDate, months).
     * Adds {@code months} calendar months to the date represented by
     * {@code excelSerial} and returns a new Excel serial date.
     */
    private static double edate(double excelSerial, int months) {
        LocalDate d = excelSerialToLocalDate(excelSerial);
        d = d.plusMonths(months);
        return localDateToExcelSerial(d);
    }

    /** Excel serial → LocalDate (Excel epoch: 1900-01-00 = day 0, treating
     *  1900 as a non-leap year per Excel's leap-year bug is irrelevant for
     *  dates after 1900-03-01; we use the standard offset of 25569 from the
     *  Unix epoch 1970-01-01). */
    private static LocalDate excelSerialToLocalDate(double serial) {
        // Excel serial 1 = 1900-01-01; but Excel treats 1900-02-29 as valid.
        // For dates >= 61 (1900-03-01) the offset from Unix epoch (day 25569)
        // works cleanly.
        long epochDay = (long) serial - 25569L; // days since 1970-01-01
        return LocalDate.ofEpochDay(epochDay);
    }

    private static double localDateToExcelSerial(LocalDate d) {
        long epochDay = d.toEpochDay();
        return (double) (epochDay + 25569L);
    }

    // ── Main computation ─────────────────────────────────────────────────────
    private void compute() {
        // ── Row 0 (Excel row 2) – seed row ───────────────────────────────────
        date[0] = Inputs.Input_IssueDate;  // =Input_IssueDate
        year[0]       = 1.0;               // =1
        yearBlank[0]  = false;
        month[0]      = 1.0;               // =1
        age[0]        = Inputs.Input_IssueAge + (year[0] - 1.0); // =Input_IssueAge+(B2-1)
        grossPrem[0]  = Inputs.Input_SchedPrem;                   // =Input_SchedPrem
        annPremYTD[0] = grossPrem[0];                             // =E2
        cumPrem[0]    = grossPrem[0];                             // =E2
        targetRem[0]  = Inputs.Input_TargetPrem;                  // =Input_TargetPrem
        premTgt[0]    = Math.min(grossPrem[0], targetRem[0]);     // =MIN(E2,H2)
        premExc[0]    = grossPrem[0] - premTgt[0];                // =E2-I2
        loadAmt[0]    = premTgt[0] * Inputs.Input_Load_Tgt
                      + premExc[0] * Inputs.Input_Load_Exc;       // =(I2*Load_Tgt)+(J2*Load_Exc)
        netPrem[0]    = grossPrem[0] - loadAmt[0];                // =E2-K2

        // ── Rows 1+ (Excel rows 3+) ───────────────────────────────────────────
        for (int i = 1; i < ROWS; i++) {
            // A: =EDATE(A_prev, 1)
            date[i] = edate(date[i - 1], 1);

            // B (year): row 3 uses >  comparison; rows 4+ use >=
            // Excel row 3: =IF(B2>Input_ProjYears,"",IF(C2=12,B2+1,B2))
            // Excel row 4+: =IF(B_prev>=Input_ProjYears,"",IF(C_prev=12,B_prev+1,B_prev))
            double prevYear  = year[i - 1];
            double prevMonth = month[i - 1];
            boolean prevBlank = yearBlank[i - 1];

            if (i == 1) {
                // row 3 formula: IF(B2 > ProjYears, "", ...)  → strict >
                if (prevYear > Inputs.Input_ProjYears) {
                    yearBlank[i] = true;
                    year[i] = 0;
                } else {
                    yearBlank[i] = false;
                    year[i] = (prevMonth == 12) ? prevYear + 1 : prevYear;
                }
            } else {
                // row 4+ formula: IF(B_prev >= ProjYears, "", ...)
                if (prevBlank || prevYear >= Inputs.Input_ProjYears) {
                    yearBlank[i] = true;
                    year[i] = 0;
                } else {
                    yearBlank[i] = false;
                    year[i] = (prevMonth == 12) ? prevYear + 1 : prevYear;
                }
            }

            // C (month): =IF(C_prev=12, 1, C_prev+1)
            month[i] = (prevMonth == 12) ? 1 : prevMonth + 1;

            // D (age): =Input_IssueAge+(B-1)   (blank year → age undefined; use 0)
            age[i] = yearBlank[i] ? 0 : Inputs.Input_IssueAge + (year[i] - 1.0);

            if (yearBlank[i]) {
                grossPrem[i]  = 0;
                annPremYTD[i] = 0;
                cumPrem[i]    = 0;
                targetRem[i]  = 0;
                premTgt[i]    = 0;
                premExc[i]    = 0;
                loadAmt[i]    = 0;
                netPrem[i]    = 0;
                continue;
            }

            // E (grossPrem): =Input_SchedPrem
            grossPrem[i] = Inputs.Input_SchedPrem;

            // F (annPremYTD): =IF(C=1, E, F_prev+E)
            annPremYTD[i] = (month[i] == 1)
                    ? grossPrem[i]
                    : annPremYTD[i - 1] + grossPrem[i];

            // G (cumPrem): =G_prev+E
            cumPrem[i] = cumPrem[i - 1] + grossPrem[i];

            // H (targetRem): =IF(C=1, Input_TargetPrem, MAX(0, Input_TargetPrem - F_prev))
            targetRem[i] = (month[i] == 1)
                    ? Inputs.Input_TargetPrem
                    : Math.max(0, Inputs.Input_TargetPrem - annPremYTD[i - 1]);

            // I (premTgt): =MIN(E, H)
            premTgt[i] = Math.min(grossPrem[i], targetRem[i]);

            // J (premExc): =E-I
            premExc[i] = grossPrem[i] - premTgt[i];

            // K (loadAmt): =(I*Load_Tgt)+(J*Load_Exc)
            loadAmt[i] = premTgt[i] * Inputs.Input_Load_Tgt
                       + premExc[i] * Inputs.Input_Load_Exc;

            // L (netPrem): =E-K
            netPrem[i] = grossPrem[i] - loadAmt[i];
        }
    }

    // ── Convenience: get net premium by 0-based index (same as col L) ────────
    public double getNetPrem(int i) { return netPrem[i]; }
}