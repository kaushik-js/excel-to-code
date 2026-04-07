public class CashValue {
    public int year;
    public int month;
    public int age;
    
    // NONG (Non-Guaranteed / Current)
    public double nongBegAv;
    public double nongNetPrem;
    public double nongCharges;
    public double nongInterest;
    public double nongEndAv;
    public double surrChg;
    public double csv; // Cash Surrender Value
    public double db;  // Death Benefit
    
    // GUAR (Guaranteed)
    public double guarBegAv;
    public double guarNetPrem;
    public double guarCharges;
    public double guarInterest;
    public double guarEndAv;

    public void calculateMonth(Inputs inputs, Premium premium, CashValue prevMonth) {
        this.year = premium.year;
        this.month = premium.month;
        this.age = premium.age;

        // 1. Set Beginning Account Values
        this.nongBegAv = (prevMonth == null) ? inputs.accumulatedCashValue : prevMonth.nongEndAv;
        this.guarBegAv = (prevMonth == null) ? inputs.accumulatedCashValue : prevMonth.guarEndAv;

        this.nongNetPrem = premium.netPrem;
        this.guarNetPrem = premium.netPrem;

        // 2. Death Benefit (Corridor testing)
        double corridorFactor = Rates.getCorridorFactor(this.age);
        this.db = Math.max(inputs.faceAmount, this.nongBegAv * corridorFactor);

        // 3. Charges (Monthly Admin + COI)
        // Net Amount at Risk (NAAR) = Death Benefit - (Beg AV + Net Prem)
        double nongNaar = Math.max(0, this.db - (this.nongBegAv + this.nongNetPrem));
        double guarNaar = Math.max(0, this.db - (this.guarBegAv + this.guarNetPrem));
        
        double coiRate = Rates.getCoiRate(this.age);
        this.nongCharges = inputs.monthlyAdminFee + (nongNaar * coiRate);
        this.guarCharges = inputs.monthlyAdminFee + (guarNaar * coiRate); // Often GUAR uses a higher guaranteed COI rate, using same here for skeleton

        // 4. Interest
        // Excel often converts annual effective to monthly using: (1 + i)^(1/12) - 1
        double nongMonthlyRate = Math.pow(1.0 + inputs.currentRateAnnual, 1.0 / 12.0) - 1.0;
        double guarMonthlyRate = Math.pow(1.0 + inputs.guaranteedRateAnnual, 1.0 / 12.0) - 1.0;

        // Value subject to interest is (Beg AV + Net Prem - Charges)
        this.nongInterest = (this.nongBegAv + this.nongNetPrem - this.nongCharges) * nongMonthlyRate;
        this.guarInterest = (this.guarBegAv + this.guarNetPrem - this.guarCharges) * guarMonthlyRate;

        // 5. End AV
        this.nongEndAv = this.nongBegAv + this.nongNetPrem - this.nongCharges + this.nongInterest;
        this.guarEndAv = this.guarBegAv + this.guarNetPrem - this.guarCharges + this.guarInterest;

        // 6. Surrender Charge & Cash Surrender Value
        this.surrChg = inputs.faceAmount * Rates.getSurrenderChargeRate(this.age);
        this.csv = Math.max(0, this.nongEndAv - this.surrChg);
    }
}