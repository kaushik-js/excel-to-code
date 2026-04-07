

public class CashValue {

    private final Inputs inputs;
    private final Premium premium;

    public CashValue(Inputs inputs, Premium premium) {
        this.inputs = inputs;
        this.premium = premium;
    }

    // ==========================================
    // Convert Annual Rate to Monthly
    // Excel equivalent: = AnnualRate / 12
    // ==========================================
    private double getMonthlyInterestRate() {
        return inputs.getCurrentRateAnnual() / 12.0;
    }

    // ==========================================
    // Monthly Mortality Rate
    // Excel equivalent: = AnnualMortalityRate / 12
    // ==========================================
    private double getMonthlyMortalityRate(int age) {
        return Rates.getMortalityRate(age) / 12.0;
    }

    // ==========================================
    // Net Amount at Risk
    // Excel: = FaceAmount - CashValue
    // ==========================================
    private double getNetAmountAtRisk(double cashValue) {
        return Math.max(0.0, inputs.getFaceAmount() - cashValue);
    }

    // ==========================================
    // Run Full Projection
    // Returns cash value at end of projection
    // ==========================================
    public double projectCashValue() {

        double cashValue = 0.0;
        int currentAge = inputs.getIssueAge();

        int totalMonths = inputs.getProjectionYears() * 12;
        double monthlyNetPremium = premium.getMonthlyNetPremium();
        double monthlyAdmin = inputs.getMonthlyAdminFee();
        double monthlyInterestRate = getMonthlyInterestRate();

        for (int month = 1; month <= totalMonths; month++) {

            // 1. Add Premium
            cashValue += monthlyNetPremium;

            // 2. Deduct Admin Charge
            cashValue -= monthlyAdmin;

            // 3. Deduct Mortality Charge
            double netAmountAtRisk = getNetAmountAtRisk(cashValue);
            double mortalityRateMonthly = getMonthlyMortalityRate(currentAge);
            double mortalityCharge = netAmountAtRisk * mortalityRateMonthly;

            cashValue -= mortalityCharge;

            // 4. Apply Interest
            cashValue += cashValue * monthlyInterestRate;

            // Prevent negative values (Excel floor behavior)
            if (cashValue < 0) {
                cashValue = 0;
            }

            // Increase age every 12 months
            if (month % 12 == 0) {
                currentAge++;
            }
        }

        return cashValue;
    }
}