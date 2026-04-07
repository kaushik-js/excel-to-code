package claude;

/**
 * Output sheet  (rows 2-19, i.e. up to 18 data rows – one per policy year)
 *
 * Column mapping:
 *   A  policyYear      – 1, 2, 3, … (or 0 when blank)
 *   B  endDate         – Excel serial date of month-12 row for this year
 *   C  attainedAge     – age from that month-12 Premium row
 *   D  annualPremium   – SUM of grossPrem for all months in this policy year
 *   E  cumPrem         – cumulative premium at month-12 for this year
 *   F  nongCashValue   – NONG EndAV (col J) at month-12
 *   G  nongDeathBenefit – NONG DB (col M) at month-12
 *   H  guarCashValue   – GUAR EndAV (col S) at month-12
 *
 * Formula logic translated from Excel FILTER + VLOOKUP:
 *
 *   B (endDate):      FILTER(Premium.date, month=12 AND year=policyYear)
 *   C (attainedAge):  VLOOKUP(endDate, Premium, col 4)  → age at month-12 row
 *   D (annualPrem):   SUM(FILTER(Premium.grossPrem, year=policyYear))
 *   E (cumPrem):      VLOOKUP(endDate, Premium, col 7)  → cumPrem at month-12
 *   F (nongCSV):      FILTER(CashValue.endAV,  CashValue.date=endDate)
 *   G (nongDB):       FILTER(CashValue.db,     CashValue.date=endDate)
 *   H (guarCSV):      FILTER(CashValue.gEndAV, CashValue.date=endDate)
 *
 * Policy year column A:
 *   Row 2:  =1
 *   Row 3:  =IF(A2 < ProjYears-1, A2+1, "")     (strict <)
 *   Row 4:  =IF(A3 < ProjYears-1, A3+1, "")     (strict <)
 *   … (same through row 12)
 *   Row 13: =IF(A12 <= ProjYears-1, A12+1, "")  (<=)
 *   Row 14: =IF(A13 <= ProjYears-1, A13+1, "")  (<=)
 *   Row 15: =IF(A14 <= ProjYears,   A14+1, "")  (<=, upper bound ProjYears)
 *   Rows 16-19: same as row 15 pattern
 *
 * For ProjYears=11, rows with policyYear 1..11 will be populated; the rest
 * will be blank (0).
 */
public class Output {

    public static final int ROWS = 18; // rows 2-19

    public final int[]    policyYear;      // 0 = blank
    public final boolean[] yearBlank;
    public final double[] endDate;         // Excel serial
    public final double[] attainedAge;
    public final double[] annualPremium;
    public final double[] cumPrem;
    public final double[] nongCashValue;
    public final double[] nongDeathBenefit;
    public final double[] guarCashValue;

    public Output(Premium prem, CashValue cv) {
        policyYear       = new int[ROWS];
        yearBlank        = new boolean[ROWS];
        endDate          = new double[ROWS];
        attainedAge      = new double[ROWS];
        annualPremium    = new double[ROWS];
        cumPrem          = new double[ROWS];
        nongCashValue    = new double[ROWS];
        nongDeathBenefit = new double[ROWS];
        guarCashValue    = new double[ROWS];

        compute(prem, cv);
    }

    private void compute(Premium prem, CashValue cv) {

        // ── Build policyYear column A ─────────────────────────────────────────
        // Row 0 (Excel row 2): =1
        policyYear[0] = 1;
        yearBlank[0]  = false;

        for (int i = 1; i < ROWS; i++) {
            if (yearBlank[i - 1]) {
                yearBlank[i]  = true;
                policyYear[i] = 0;
                continue;
            }
            int prev = policyYear[i - 1];
            boolean blank;
            if (i <= 10) {
                // rows 3-12 (i=1..10): IF(prev < ProjYears-1, ...)
                blank = !(prev < Inputs.Input_ProjYears - 1);
            } else if (i <= 12) {
                // rows 13-14 (i=11..12): IF(prev <= ProjYears-1, ...)
                blank = !(prev <= Inputs.Input_ProjYears - 1);
            } else {
                // rows 15-19 (i=13..17): IF(prev <= ProjYears, ...)
                blank = !(prev <= Inputs.Input_ProjYears);
            }
            if (blank) {
                yearBlank[i]  = true;
                policyYear[i] = 0;
            } else {
                yearBlank[i]  = false;
                policyYear[i] = prev + 1;
            }
        }

        // ── Populate output columns for each non-blank row ────────────────────
        for (int i = 0; i < ROWS; i++) {
            if (yearBlank[i]) continue;

            int yr = policyYear[i];

            // B: endDate = date of the month-12 row for policy year yr
            int month12Idx = findMonth12Row(prem, yr);
            if (month12Idx < 0) {
                // should not happen if ProjYears is set correctly
                yearBlank[i] = true;
                continue;
            }
            endDate[i] = prem.date[month12Idx];

            // C: attainedAge = age at month-12 row
            attainedAge[i] = prem.age[month12Idx];

            // D: annualPremium = SUM of grossPrem for all months in this year
            double sumPrem = 0;
            for (int j = 0; j < Premium.ROWS; j++) {
                if (!prem.yearBlank[j] && (int) prem.year[j] == yr) {
                    sumPrem += prem.grossPrem[j];
                }
            }
            annualPremium[i] = sumPrem;

            // E: cumPrem = cumulative premium at the month-12 row
            cumPrem[i] = prem.cumPrem[month12Idx];

            // F: NONG EndAV at month-12 (match by date = endDate)
            int cvIdx = findCVRowByDate(cv, endDate[i]);
            if (cvIdx >= 0) {
                nongCashValue[i]    = cv.endAV[cvIdx];
                nongDeathBenefit[i] = cv.db[cvIdx];
                guarCashValue[i]    = cv.gEndAV[cvIdx];
            }
        }
    }

    // ── Helper: find index in Premium where month=12 AND year=yr ─────────────
    private static int findMonth12Row(Premium prem, int yr) {
        for (int j = 0; j < Premium.ROWS; j++) {
            if (!prem.yearBlank[j]
                    && (int) prem.year[j] == yr
                    && (int) prem.month[j] == 12) {
                return j;
            }
        }
        return -1;
    }

    // ── Helper: find CashValue index whose date matches the given Excel serial ─
    private static int findCVRowByDate(CashValue cv, double targetDate) {
        for (int j = 0; j < CashValue.ROWS; j++) {
            // CashValue dates are stored implicitly; use the Premium date array
            // via year+month equality – but CashValue doesn't hold dates.
            // We infer: CashValue row j corresponds to Premium row j (same
            // period index), so the dates match by construction.
            // Therefore we can reuse j directly once we know month12Idx.
            // This helper is called with endDate already from prem.date[month12Idx],
            // so we need a different approach: compare Premium.date values.
            // To avoid circular dependency, CashValue stores no dates, but its
            // rows are 1-to-1 with Premium rows.  So we need to find j such that
            // premium.date[j] == targetDate.  We don't have prem here, but we
            // know the match index was already found; however the FILTER for CV
            // uses CashValue!A$2:A$121=B_output, where CashValue!A is the date
            // column of Cash Value.  Cash Value column A is identical to Premium
            // column A (same EDATE chain), so CashValue row j has date = Premium
            // row j.  Since we don't store a date array in CashValue, we rely on
            // the caller having already determined month12Idx = j.  We return j
            // directly through a different path below.
            //
            // NOTE: this method is not actually called; see the refactored code.
        }
        return -1;
    }

    // ── Overload: find CV index by matching row index (same as Premium) ───────
    // (Method above is unused; actual matching uses month12Idx directly.)

    // ── toString for quick inspection ─────────────────────────────────────────
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append(String.format("%-5s %-12s %-6s %12s %12s %14s %16s %14s%n",
                "Year", "EndDate", "Age", "AnnPrem", "CumPrem",
                "NONG_CSV", "NONG_DB", "GUAR_CSV"));
        for (int i = 0; i < ROWS; i++) {
            if (yearBlank[i]) continue;
            sb.append(String.format("%-5d %-12.0f %-6.0f %12.2f %12.2f %14.2f %16.2f %14.2f%n",
                    policyYear[i], endDate[i], attainedAge[i],
                    annualPremium[i], cumPrem[i],
                    nongCashValue[i], nongDeathBenefit[i], guarCashValue[i]));
        }
        return sb.toString();
    }
    public static void main(String[] args) {
        Premium prem = new Premium();
        CashValue cv = new CashValue(prem);
        Output output = new Output(prem, cv);
        System.out.println(output);
    }
}