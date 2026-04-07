public class Premium {
    public int year;
    public int month;
    public int age;
    public double grossPrem;
    public double annPremYTD;
    public double cumPrem;
    public double targetRem;
    public double premTgt;
    public double premExc;
    public double loadAmt;
    public double netPrem;

    public void calculatePremium(Inputs inputs, Premium prevMonth, double currentGrossPrem, int currentYear, int currentMonth, int currentAge) {
        this.year = currentYear;
        this.month = currentMonth;
        this.age = currentAge;
        this.grossPrem = currentGrossPrem;

        double prevAnnPremYTD = (prevMonth != null && prevMonth.year == this.year) ? prevMonth.annPremYTD : 0.0;
        double prevCumPrem = (prevMonth != null) ? prevMonth.cumPrem : 0.0;
        double prevTargetRem = (prevMonth != null && prevMonth.year == this.year) ? prevMonth.targetRem : inputs.targetPremiumAnnual;

        this.annPremYTD = prevAnnPremYTD + this.grossPrem;
        this.cumPrem = prevCumPrem + this.grossPrem;

        // Split premium into Up-to-Target and Excess
        if (this.grossPrem <= prevTargetRem) {
            this.premTgt = this.grossPrem;
            this.premExc = 0.0;
        } else {
            this.premTgt = Math.max(0, prevTargetRem);
            this.premExc = this.grossPrem - this.premTgt;
        }

        this.targetRem = Math.max(0, prevTargetRem - this.grossPrem);
        this.loadAmt = (this.premTgt * inputs.premLoadUpToTarget) + (this.premExc * inputs.premLoadExcess);
        this.netPrem = this.grossPrem - this.loadAmt;
    }
}