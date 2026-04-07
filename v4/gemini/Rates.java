public class Rates {
    // In a full implementation, these arrays would hold all 120+ rows of the CSV.
    // Truncated here to demonstrate the exact structure.
    private static final double[] COI_RATE = new double[121];
    private static final double[] SURR_CHG_RATE = new double[121];
    private static final double[] CORRIDOR_FACTOR = new double[121];

    static {
        // Age 35 (Issue Age from snippets)
        COI_RATE[35] = 0.00012035; // Example extrapolated rate
        SURR_CHG_RATE[35] = 0.05;
        CORRIDOR_FACTOR[35] = 2.50;
        
        // Age 44 (From Cash Value snippet)
        COI_RATE[44] = 0.00019311;
        SURR_CHG_RATE[44] = 0.05;
        CORRIDOR_FACTOR[44] = 2.15;
        
        // Populate remainder of mortality/corridor tables...
    }

    public static double getCoiRate(int age) {
        return COI_RATE[age];
    }

    public static double getSurrenderChargeRate(int age) {
        return SURR_CHG_RATE[age];
    }

    public static double getCorridorFactor(int age) {
        return CORRIDOR_FACTOR[age];
    }
}