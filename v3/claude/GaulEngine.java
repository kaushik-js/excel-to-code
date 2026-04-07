package claude;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

/**
 * GaulEngine - Java translation of gaulEngine_v2.xlsx
 *
 * Models a Universal Life insurance policy with:
 *   - Premium sheet: monthly premium allocation (target vs. excess, loads)
 *   - Cash Value sheet: non-guaranteed (NONG) and guaranteed (GUAR) account values
 *   - Output sheet: annual summary
 *
 * All logic is an exact translation of the Excel formulas; no randomness or
 * system-time dependencies are introduced.
 */
public class GaulEngine {

    // =========================================================================
    // Inputs (Inputs sheet, column B)
    // =========================================================================
    static final double INPUT_FACE_AMT      = 100_000.0;       // B2
    static final int    INPUT_ISSUE_AGE     = 35;              // B3
    // B4 Issue Date: 2009-03-27  – used only as a display anchor; month offsets
    // are computed via the EDATE logic (add N months to issue date).
    static final int    ISSUE_YEAR          = 2009;
    static final int    ISSUE_MONTH         = 3;   // 1-based
    static final int    ISSUE_DAY           = 27;
    static final int    INPUT_PROJ_YEARS    = 11;              // B8
    static final double INPUT_SCHED_PREM    = 50.0;            // B10
    static final double INPUT_TARGET_PREM   = 630.71;          // B11
    // B12 / B13 – Guideline limits (present in sheet, not used in formulas)
    @SuppressWarnings("unused")
    static final double INPUT_GL_ANNUAL     = 1_049.05;
    @SuppressWarnings("unused")
    static final double INPUT_GL_SINGLE     = 11_801.53;
    static final double INPUT_LOAD_TGT      = 0.07;            // B15
    static final double INPUT_LOAD_EXC      = 0.05;            // B16
    static final double ACV                 = 11_410.54;       // B17 (Accumulated Cash Value)
    static final double INPUT_ADMIN_FEE     = 5.0;             // B18
    static final double INPUT_INT_CUR       = 0.0352;          // B20
    static final double INPUT_INT_GUAR      = 0.03;            // B21

    // Calculated monthly rates (named formulas in workbook)
    static final double CALC_MTHLY_RATE_CUR  = Math.pow(1.0 + INPUT_INT_CUR,  1.0 / 12.0) - 1.0;
    static final double CALC_MTHLY_RATE_GUAR = Math.pow(1.0 + INPUT_INT_GUAR, 1.0 / 12.0) - 1.0;

    // Total rows = projYears * 12 months + header row 1; data rows 2..TOTAL_ROWS
    static final int TOTAL_MONTHS = INPUT_PROJ_YEARS * 12;   // 132 months for 11 years

    // =========================================================================
    // Rates table – Rates sheet A2:D122  (ages 0–120)
    // COI_Rate, SurrChg_Rate, Corridor_Factor indexed by age (index = age)
    // =========================================================================
    // Pre-built from the sheet data
    static final double[] COI_RATE       = new double[122]; // indices 0..121
    static final double[] SURR_CHG_RATE  = new double[122];
    static final double[] CORRIDOR_FACTOR = new double[122];

    static {
        // Ages 0–40: SurrChg = 0.05, Corridor = 2.5, COI starts at 1.5e-5 and
        // compounds by factor 1.065 each year (derived from the sheet data).
        // Ages 41–95: Corridor decreases from 2.472727... to 1.0
        // Ages 96–120+: Corridor = 1.0
        // All SurrChg_Rate values are 0.05 throughout.
        //
        // Rather than re-derive the formula, we embed the exact values from the sheet.

        double[] ages = {
            0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,
            20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,
            40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,
            60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,
            80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,
            100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,
            115,116,117,118,119,120,121
        };
        double[] coi = {
            1.5e-05,1.5975e-05,1.7013375e-05,1.8119244375e-05,1.9296995259375e-05,
            2.055129995123437e-05,2.18871344480646e-05,2.33097981871888e-05,
            2.482493506935607e-05,2.643855584886422e-05,2.815706197904039e-05,
            2.998727100767801e-05,3.193644362317708e-05,3.401231245868359e-05,
            3.622311276849802e-05,3.857761509845039e-05,4.108516007984966e-05,
            4.375569548503989e-05,4.659981569156749e-05,4.962880371151936e-05,
            5.285467595276812e-05,5.629022988969805e-05,5.994909483252842e-05,
            6.384578599664275e-05,6.799576208642454e-05,7.241548662204213e-05,
            7.712249325247486e-05,8.213545531388573e-05,8.747425990928829e-05,
            9.316008680339203e-05,9.921549244561249e-05,1.056644994545773e-04,
            1.125326919191248e-04,1.198473168938679e-04,1.276373924919693e-04,
            1.359338230039474e-04,1.447695214992039e-04,1.541795403966522e-04,
            1.642012105224345e-04,1.748742892063928e-04,1.862411180048083e-04,
            1.983467906751209e-04,2.112393320690037e-04,2.249698886534889e-04,
            2.395929314159657e-04,2.551664719580034e-04,2.717522926352736e-04,
            2.894161916565664e-04,3.082282441142432e-04,3.28263079981669e-04,
            3.496001801804775e-04,3.723241918922085e-04,3.96525264365202e-04,
            4.222994065489401e-04,4.497488679746212e-04,4.789825443929715e-04,
            5.101164097785147e-04,5.432739764141181e-04,5.785867848810357e-04,
            6.16194925898303e-04,6.562475960816927e-04,6.989036898270027e-04,
            7.443324296657578e-04,7.927140375940321e-04,8.442404500376441e-04,
            8.991160792900909e-04,9.575586244439468e-04,1.019799935032803e-03,
            1.086086930809936e-03,1.156682581312581e-03,1.231866949097899e-03,
            1.311938300789262e-03,1.397214290340564e-03,1.488033219212701e-03,
            1.584755378461526e-03,1.687764478061525e-03,1.797469169135524e-03,
            1.914304665129334e-03,2.03873446836274e-03,2.171252208806318e-03,
            2.312383602378729e-03,2.462688536533346e-03,2.622763291408013e-03,
            2.793242905349534e-03,2.974803694197254e-03,3.168165934320075e-03,
            3.37409672005088e-03,3.593413006854187e-03,3.826984852299709e-03,
            4.075738867699189e-03,4.340661894099637e-03,4.622804917216112e-03,
            4.92328723683516e-03,5.243300907229445e-03,5.584115466199358e-03,
            5.947082971502316e-03,6.333643364649967e-03,6.745330183352214e-03,
            7.183776645270108e-03,7.650722127212664e-03,8.148019065481488e-03,
            8.677640304737784e-03,9.24168692454574e-03,9.842396574641211e-03,
            1.048215235199289e-02,1.116349225487243e-02,1.188911925143913e-02,
            1.266191200278268e-02,1.348493628296355e-02,1.436145714135618e-02,
            1.529495185554433e-02,1.628912372615471e-02,1.734791676835477e-02,
            1.847553135829783e-02,1.967644089658718e-02,2.095540955486535e-02,
            2.23175111759316e-02,2.376814940236715e-02,2.531307911352101e-02,
            2.695842925589988e-02,2.871072715753337e-02,3.057692442277303e-02
        };
        double[] corr = {
            2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,
            2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,
            2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,2.5,
            2.472727272727273,2.445454545454545,2.418181818181818,2.390909090909091,
            2.363636363636364,2.336363636363636,2.309090909090909,2.281818181818182,
            2.254545454545454,2.227272727272728,2.2,2.172727272727273,
            2.145454545454546,2.118181818181818,2.090909090909091,2.063636363636363,
            2.036363636363636,2.009090909090909,1.981818181818182,1.954545454545455,
            1.927272727272727,1.9,1.872727272727273,1.845454545454545,
            1.818181818181818,1.790909090909091,1.763636363636364,1.736363636363636,
            1.709090909090909,1.681818181818182,1.654545454545455,1.627272727272727,
            1.6,1.572727272727273,1.545454545454545,1.518181818181818,
            1.490909090909091,1.463636363636364,1.436363636363637,1.409090909090909,
            1.381818181818182,1.354545454545454,1.327272727272727,1.3,
            1.272727272727273,1.245454545454546,1.218181818181818,1.190909090909091,
            1.163636363636364,1.136363636363636,1.109090909090909,1.081818181818182,
            1.054545454545455,1.027272727272727,
            1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,
            1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0
        };

        for (int i = 0; i < ages.length; i++) {
            int age = (int) ages[i];
            COI_RATE[age]        = coi[i];
            SURR_CHG_RATE[age]   = 0.05;
            CORRIDOR_FACTOR[age] = corr[i];
        }
    }

    // =========================================================================
    // EDATE helper: add months to a date represented as (year, month, day)
    // Returns [year, month, day].  Day is clamped to end-of-month.
    // =========================================================================
    static int[] edate(int year, int month, int day, int addMonths) {
        int totalMonths = (year - 1) * 12 + (month - 1) + addMonths;
        int y = totalMonths / 12 + 1;
        int m = totalMonths % 12 + 1;
        // Clamp day to valid range for month
        int maxDay = daysInMonth(y, m);
        int d = Math.min(day, maxDay);
        return new int[]{y, m, d};
    }

    static int daysInMonth(int year, int month) {
        switch (month) {
            case 1: case 3: case 5: case 7: case 8: case 10: case 12: return 31;
            case 4: case 6: case 9: case 11: return 30;
            case 2: return (isLeap(year) ? 29 : 28);
        }
        return 30;
    }

    static boolean isLeap(int year) {
        return (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
    }

    // =========================================================================
    // Row data containers
    // =========================================================================
    record PremiumRow(
        int[]  date,   // [year,month,day]
        int    yr,     // policy year (1-based), -1 = past end
        int    mon,    // policy month (1-12)
        int    age,
        double grossPrem,
        double annPremYTD,
        double cumPrem,
        double targetRem,
        double premTgt,
        double premExc,
        double loadAmt,
        double netPrem
    ) {}

    record CvRow(
        int[]  date,
        int    yr,
        int    mon,
        int    age,
        // NONG
        double begAV,
        double netPrem,
        double charges,
        double interest,
        double endAV,
        double surrChg,
        double csv,
        double db,
        // GUAR
        double gBegAV,
        double gNetPrem,
        double gCharges,
        double gInterest,
        double gEndAV
    ) {}

    // =========================================================================
    // Main
    // =========================================================================
    public static void main(String[] args) {

        // --- Build Premium rows ---
        PremiumRow[] prem = buildPremiumRows();

        // --- Build Cash Value rows ---
        CvRow[] cv = buildCvRows(prem);

        // --- Print Output sheet ---
        printOutput(prem, cv);

        // Optionally print full monthly detail
        // printMonthlyDetail(prem, cv);
    }

    // =========================================================================
    // Build Premium rows (Premium sheet formulas)
    // =========================================================================
    static PremiumRow[] buildPremiumRows() {
        PremiumRow[] rows = new PremiumRow[TOTAL_MONTHS];

        for (int i = 0; i < TOTAL_MONTHS; i++) {
            int[]  date;
            int    yr, mon, age;
            double grossPrem, annPremYTD, cumPrem, targetRem, premTgt, premExc, loadAmt, netPrem;

            if (i == 0) {
                // Row 2 (first data row)
                date = new int[]{ISSUE_YEAR, ISSUE_MONTH, ISSUE_DAY};
                yr   = 1;
                mon  = 1;
                age  = INPUT_ISSUE_AGE + (yr - 1);
                grossPrem   = INPUT_SCHED_PREM;
                annPremYTD  = grossPrem;
                cumPrem     = grossPrem;
                targetRem   = INPUT_TARGET_PREM;
                premTgt     = Math.min(grossPrem, targetRem);
                premExc     = grossPrem - premTgt;
                loadAmt     = premTgt * INPUT_LOAD_TGT + premExc * INPUT_LOAD_EXC;
                netPrem     = grossPrem - loadAmt;
            } else {
                PremiumRow prev = rows[i - 1];

                // Date = EDATE(prev date, 1)
                date = edate(prev.date()[0], prev.date()[1], prev.date()[2], 1);

                // Year column – mirrors the two-variant Excel formula:
                //   Row 3: IF(B2>projYears,"", IF(C2=12,B2+1,B2))
                //   Row 4+: IF(B3>=projYears,"", IF(C3=12,B3+1,B3))
                boolean prevYrIsBlank = (prev.yr() < 0);
                if (i == 1) {
                    // row 3: condition B2 > projYears
                    yr = (prev.yr() > INPUT_PROJ_YEARS) ? -1
                       : (prev.mon() == 12 ? prev.yr() + 1 : prev.yr());
                } else {
                    // row 4+: condition B_prev >= projYears
                    yr = (prevYrIsBlank || prev.yr() >= INPUT_PROJ_YEARS) ? -1
                       : (prev.mon() == 12 ? prev.yr() + 1 : prev.yr());
                }

                mon = (prev.mon() == 12) ? 1 : prev.mon() + 1;
                age = (yr < 0) ? (INPUT_ISSUE_AGE - 1) : (INPUT_ISSUE_AGE + (yr - 1));

                boolean blank = (yr < 0);
                grossPrem   = blank ? 0 : INPUT_SCHED_PREM;
                annPremYTD  = blank ? 0 : (mon == 1 ? grossPrem : prev.annPremYTD() + grossPrem);
                cumPrem     = blank ? 0 : prev.cumPrem() + grossPrem;
                targetRem   = blank ? 0 : (mon == 1 ? INPUT_TARGET_PREM
                                          : Math.max(0, INPUT_TARGET_PREM - prev.annPremYTD()));
                premTgt     = blank ? 0 : Math.min(grossPrem, targetRem);
                premExc     = blank ? 0 : grossPrem - premTgt;
                loadAmt     = blank ? 0 : premTgt * INPUT_LOAD_TGT + premExc * INPUT_LOAD_EXC;
                netPrem     = blank ? 0 : grossPrem - loadAmt;
            }

            rows[i] = new PremiumRow(date, yr, mon, age, grossPrem, annPremYTD, cumPrem,
                                     targetRem, premTgt, premExc, loadAmt, netPrem);
        }
        return rows;
    }

    // =========================================================================
    // Build Cash Value rows (Cash Value sheet formulas)
    // =========================================================================
    static CvRow[] buildCvRows(PremiumRow[] prem) {
        CvRow[] rows = new CvRow[TOTAL_MONTHS];

        for (int i = 0; i < TOTAL_MONTHS; i++) {
            PremiumRow p = prem[i];
            int[]  date  = p.date();
            int    yr    = p.yr();
            int    mon   = p.mon();
            int    age   = p.age();
            boolean blank = (yr < 0);

            // NONG
            double begAV   = (i == 0) ? ACV : (blank ? 0 : rows[i-1].endAV());
            double netPrem = p.netPrem();           // comes from Premium col L
            double coiRate = lookupCOI(age);
            double charges, interest, endAV, surrChg, csv, db;
            if (blank) {
                charges = 0; interest = 0; endAV = 0; surrChg = 0; csv = 0; db = 0;
            } else {
                charges  = coiRate * Math.max(0, INPUT_FACE_AMT - (begAV + netPrem)) + INPUT_ADMIN_FEE;
                interest = (begAV + netPrem - charges) * CALC_MTHLY_RATE_CUR;
                endAV    = Math.max(0, begAV + netPrem - charges + interest);
                surrChg  = lookupSurrChg(age) * INPUT_FACE_AMT;
                csv      = Math.max(0, endAV - surrChg);
                db       = Math.max(INPUT_FACE_AMT, endAV * lookupCorridorFactor(age));
            }

            // GUAR – always starts from ACV (O column = ACV in every row per formula)
            double gBegAV   = ACV;   // O column: =ACV (hardcoded to ACV every row)
            double gNetPrem = netPrem;
            double gCharges, gInterest, gEndAV;
            if (blank) {
                gCharges = 0; gInterest = 0; gEndAV = 0;
            } else {
                double guarCoiRate = coiRate * 1.5;
                gCharges  = guarCoiRate * Math.max(0, INPUT_FACE_AMT - (gBegAV + gNetPrem)) + INPUT_ADMIN_FEE;
                gInterest = (gBegAV + gNetPrem - gCharges) * CALC_MTHLY_RATE_GUAR;
                gEndAV    = Math.max(0, gBegAV + gNetPrem - gCharges + gInterest);
            }

            rows[i] = new CvRow(date, yr, mon, age,
                                 begAV, netPrem, charges, interest, endAV, surrChg, csv, db,
                                 gBegAV, gNetPrem, gCharges, gInterest, gEndAV);
        }
        return rows;
    }

    // =========================================================================
    // Rate table lookups (VLOOKUP exact match = FALSE)
    // =========================================================================
    static double lookupCOI(int age) {
        if (age < 0 || age > 121) return 0;
        return COI_RATE[age];
    }
    static double lookupSurrChg(int age) {
        if (age < 0 || age > 121) return 0;
        return SURR_CHG_RATE[age];
    }
    static double lookupCorridorFactor(int age) {
        if (age < 0 || age > 121) return 1.0;
        return CORRIDOR_FACTOR[age];
    }

    // =========================================================================
    // Output sheet – annual summary
    // Column lookups replicate VLOOKUP on end-of-year rows (month 12 of each yr)
    // =========================================================================
    static void printOutput(PremiumRow[] prem, CvRow[] cv) {
        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("MM/dd/yyyy");

        System.out.println("=== OUTPUT (Annual Summary) ===");
        System.out.printf("%-6s  %-12s  %-5s  %14s  %16s  %16s  %18s  %14s%n",
            "Year", "End Date", "Age", "Annual Prem", "Cum Prem",
            "NONG CSV", "NONG Death Benefit", "GUAR End AV");
        System.out.println("-".repeat(105));

        // Output rows 2..up to projYears.
        // For each policy year Y (1-based), find the last month row of that year.
        // The Output B column uses an array formula to find the date of month 12
        // in year Y; C/D/E/F/G/H use VLOOKUP against Premium/CashValue tables.
        for (int y = 1; y <= INPUT_PROJ_YEARS; y++) {
            // find index of month=12 for year=y (or last month of last year)
            int idx = -1;
            for (int i = 0; i < TOTAL_MONTHS; i++) {
                if (prem[i].yr() == y && prem[i].mon() == 12) { idx = i; break; }
            }
            if (idx < 0) {
                // year y has no month-12 row — use last valid row
                for (int i = TOTAL_MONTHS - 1; i >= 0; i--) {
                    if (prem[i].yr() == y) { idx = i; break; }
                }
            }
            if (idx < 0) break;

            PremiumRow pr = prem[idx];
            CvRow cr      = cv[idx];

            int[] d = pr.date();
            LocalDate endDate = LocalDate.of(d[0], d[1], d[2]);
            System.out.printf("%-6d  %-12s  %-5d  %14.2f  %16.2f  %16.2f  %18.2f  %14.2f%n",
                y,
                endDate.format(fmt),
                pr.age(),
                pr.annPremYTD(),
                pr.cumPrem(),
                cr.csv(),
                cr.db(),
                cr.gEndAV());
        }
        System.out.println();
    }

    // =========================================================================
    // Optional: full monthly detail
    // =========================================================================
    @SuppressWarnings("unused")
    static void printMonthlyDetail(PremiumRow[] prem, CvRow[] cv) {
        System.out.println("=== PREMIUM DETAIL ===");
        System.out.printf("%-4s %-3s %-3s %-5s %8s %12s %12s %12s %12s %12s %12s %12s%n",
            "Yr","Mo","Age","GrPrm","AnnYTD","CumPrem","TgtRem","PremTgt","PremExc","LoadAmt","NetPrem");
        for (PremiumRow r : prem) {
            if (r.yr() < 0) break;
            System.out.printf("%-4d %-3d %-3d %8.2f %12.2f %12.2f %12.2f %12.2f %12.2f %12.4f %12.2f%n",
                r.yr(), r.mon(), r.age(), r.grossPrem(), r.annPremYTD(), r.cumPrem(),
                r.targetRem(), r.premTgt(), r.premExc(), r.loadAmt(), r.netPrem());
        }

        System.out.println("\n=== CASH VALUE DETAIL ===");
        System.out.printf("%-4s %-3s %-3s %12s %10s %10s %10s %10s %10s %10s %10s | %12s %10s %10s %10s %10s%n",
            "Yr","Mo","Age","BegAV","NetPrem","Charges","Interest","EndAV","SurrChg","CSV","DB",
            "G-BegAV","G-NetPrm","G-Chgs","G-Int","G-EndAV");
        for (CvRow r : cv) {
            if (r.yr() < 0) break;
            System.out.printf("%-4d %-3d %-3d %12.2f %10.2f %10.4f %10.4f %10.2f %10.2f %10.2f %10.2f | %12.2f %10.2f %10.4f %10.4f %10.2f%n",
                r.yr(), r.mon(), r.age(),
                r.begAV(), r.netPrem(), r.charges(), r.interest(), r.endAV(),
                r.surrChg(), r.csv(), r.db(),
                r.gBegAV(), r.gNetPrem(), r.gCharges(), r.gInterest(), r.gEndAV());
        }
    }
}