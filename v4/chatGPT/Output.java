public class Output {

    private final Inputs inputs;
    private final Premium premium;

    public Output(Inputs inputs, Premium premium) {
        this.inputs = inputs;
        this.premium = premium;
    }

    public void generateProjectionReport() {

        int projectionYears = inputs.getProjectionYears();
        int currentAge = inputs.getIssueAge();

        double cashValue = 0.0;

        double monthlyNetPremium = premium.getMonthlyNetPremium();
        double monthlyAdmin = inputs.getMonthlyAdminFee();
        double monthlyInterestRate = inputs.getCurrentRateAnnual() / 12.0;

        System.out.println("YEAR | AGE | ANNUAL PREMIUM | END CASH VALUE");
        System.out.println("------------------------------------------------");

        for (int year = 1; year <= projectionYears; year++) {

            for (int month = 1; month <= 12; month++) {

                // Add Premium
                cashValue += monthlyNetPremium;

                // Admin Charge
                cashValue -= monthlyAdmin;

                // Mortality
                double netAmountAtRisk = Math.max(0.0, inputs.getFaceAmount() - cashValue);
                double mortalityRateMonthly = Rates.getMortalityRate(currentAge) / 12.0;
                double mortalityCharge = netAmountAtRisk * mortalityRateMonthly;

                cashValue -= mortalityCharge;

                // Interest
                cashValue += cashValue * monthlyInterestRate;

                if (cashValue < 0) {
                    cashValue = 0;
                }
            }

            // Print End-of-Year Summary
            System.out.printf(
                "%4d | %3d | %14.2f | %15.2f%n",
                year,
                currentAge,
                premium.getAnnualPremium(),
                cashValue
            );

            currentAge++;
        }

        System.out.println("------------------------------------------------");
        System.out.println("Exceeds Guideline Annual Limit: " + premium.exceedsGuidelineAnnualLimit());
        System.out.println("Exceeds Guideline Single Limit: " + premium.exceedsGuidelineSingleLimit());
    }

    public static void main(String[] args) {
        Inputs inputs = new Inputs();
        Premium premium = new Premium(inputs);
        Output output = new Output(inputs, premium);
        output.generateProjectionReport();
    }
}