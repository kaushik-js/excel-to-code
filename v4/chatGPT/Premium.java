public class Premium {

    private final Inputs inputs;

    public Premium(Inputs inputs) {
        this.inputs = inputs;
    }

    // ==========================================
    // Annual Premium
    // Excel: = ScheduledMonthlyPremium * 12
    // ==========================================
    public double getAnnualPremium() {
        return inputs.getScheduledMonthlyPremium() * 12.0;
    }

    // ==========================================
    // Premium Load Calculation
    //
    // Excel Equivalent:
    // =IF(AnnualPremium <= TargetPremiumAnnual,
    //      AnnualPremium * PremiumLoadUpToTarget,
    //      (TargetPremiumAnnual * PremiumLoadUpToTarget)
    //      + ((AnnualPremium - TargetPremiumAnnual) * PremiumLoadExcess)
    // )
    // ==========================================
    public double getPremiumLoad() {

        double annualPremium = getAnnualPremium();
        double target = inputs.getTargetPremiumAnnual();
        double loadUpToTarget = inputs.getPremiumLoadUpToTarget();
        double loadExcess = inputs.getPremiumLoadExcess();

        if (annualPremium <= target) {
            return annualPremium * loadUpToTarget;
        } else {
            double loadWithinTarget = target * loadUpToTarget;
            double excessLoad = (annualPremium - target) * loadExcess;
            return loadWithinTarget + excessLoad;
        }
    }

    // ==========================================
    // Net Annual Premium
    // Excel: = AnnualPremium - PremiumLoad
    // ==========================================
    public double getNetAnnualPremium() {
        return getAnnualPremium() - getPremiumLoad();
    }

    // ==========================================
    // Monthly Net Premium
    // Excel: = NetAnnualPremium / 12
    // ==========================================
    public double getMonthlyNetPremium() {
        return getNetAnnualPremium() / 12.0;
    }

    // ==========================================
    // Guideline Compliance Checks
    // ==========================================

    // Excel: = AnnualPremium > GuidelineAnnualLimit
    public boolean exceedsGuidelineAnnualLimit() {
        return getAnnualPremium() > inputs.getGuidelineAnnualLimit();
    }

    // Excel: = AnnualPremium > GuidelineSingleLimit
    public boolean exceedsGuidelineSingleLimit() {
        return getAnnualPremium() > inputs.getGuidelineSingleLimit();
    }

}