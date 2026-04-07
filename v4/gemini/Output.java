public class Output {
    public int policyYear;
    public int attainedAge;
    public double annualPremium;
    public double cumulativePremium;
    public double nongCashValue;
    public double nongDeathBenefit;
    public double guarCashValue;

    public static Output extractAnnualSummary(CashValue endOfYearMonth, Premium endOfYearPremium) {
        Output out = new Output();
        out.policyYear = endOfYearMonth.year;
        out.attainedAge = endOfYearMonth.age;
        out.annualPremium = endOfYearPremium.annPremYTD;
        out.cumulativePremium = endOfYearPremium.cumPrem;
        out.nongCashValue = endOfYearMonth.nongEndAv;
        out.nongDeathBenefit = endOfYearMonth.db;
        out.guarCashValue = endOfYearMonth.guarEndAv;
        return out;
    }
    public static void main(String[] args) {
        Premium prem = new Premium();
        CashValue cv = new CashValue();
        System.out.println(extractAnnualSummary(cv, prem));
    }
}