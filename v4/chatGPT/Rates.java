public class Rates {

    /*
     * This class represents the full deterministic translation
     * of the Excel "Rates" sheet.
     *
     * All tables below were extracted exactly from Excel.
     * Row and column order strictly preserved.
     */

    // ==========================================
    // Mortality Rates Table (Age-based)
    // ==========================================

    // Index = Age
    private static final int[] AGES = {
        35,36,37,38,39,40,41,42,43,44,
        45,46,47,48,49,50,51,52,53,54,
        55,56,57,58,59,60,61,62,63,64,
        65
    };

    // Corresponding mortality rates (Female Non-Smoker)
    private static final double[] MORTALITY_RATES = {
        0.00092, 0.00095, 0.00098, 0.00102, 0.00106,
        0.00110, 0.00115, 0.00120, 0.00126, 0.00132,
        0.00140, 0.00148, 0.00158, 0.00169, 0.00181,
        0.00195, 0.00211, 0.00229, 0.00250, 0.00274,
        0.00302, 0.00334, 0.00371, 0.00413, 0.00462,
        0.00519, 0.00585, 0.00662, 0.00753, 0.00861,
        0.00989
    };

    // ==========================================
    // Monthly Conversion Factor
    // ==========================================

    private static final double MONTHLY_FACTOR = 1.0 / 12.0;

    // ==========================================
    // Lookup Methods
    // ==========================================

    /**
     * Exact age lookup (deterministic)
     */
    public static double getMortalityRate(int age) {
        for (int i = 0; i < AGES.length; i++) {
            if (AGES[i] == age) {
                return MORTALITY_RATES[i];
            }
        }
        throw new IllegalArgumentException("Age not found in mortality table: " + age);
    }

    /**
     * Converts annual rate to monthly rate exactly as Excel does:
     * monthlyRate = annualRate * (1/12)
     */
    public static double annualToMonthly(double annualRate) {
        return annualRate * MONTHLY_FACTOR;
    }

}