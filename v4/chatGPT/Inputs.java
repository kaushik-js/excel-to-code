import java.time.LocalDate;

public class Inputs {

    // Policy Setup
    private final double faceAmount = 100000.0;
    private final int issueAge = 35;
    private final LocalDate issueDate = LocalDate.of(2009, 3, 27);
    private final String gender = "Female";
    private final String policyClass = "Non-Smoker";

    // Projection
    private final int projectionYears = 11;

    // Premiums
    private final double scheduledMonthlyPremium = 50.0;
    private final double targetPremiumAnnual = 630.71;
    private final double guidelineAnnualLimit = 1049.05;
    private final double guidelineSingleLimit = 11801.53;

    // Loads
    private final double premiumLoadUpToTarget = 0.07;
    private final double premiumLoadExcess = 0.05;

    // Charges
    private final double monthlyAdminFee = 5.0;

    // Rates
    private final double currentRateAnnual = 0.0352;
    private final double guaranteedRateAnnual = 0.03;

    // ======================
    // Getters (ALL REQUIRED)
    // ======================

    public double getFaceAmount() { return faceAmount; }
    public int getIssueAge() { return issueAge; }
    public LocalDate getIssueDate() { return issueDate; }
    public String getGender() { return gender; }
    public String getPolicyClass() { return policyClass; }

    public int getProjectionYears() { return projectionYears; }

    public double getScheduledMonthlyPremium() { return scheduledMonthlyPremium; }
    public double getTargetPremiumAnnual() { return targetPremiumAnnual; }
    public double getGuidelineAnnualLimit() { return guidelineAnnualLimit; }
    public double getGuidelineSingleLimit() { return guidelineSingleLimit; }

    public double getPremiumLoadUpToTarget() { return premiumLoadUpToTarget; }
    public double getPremiumLoadExcess() { return premiumLoadExcess; }

    public double getMonthlyAdminFee() { return monthlyAdminFee; }

    public double getCurrentRateAnnual() { return currentRateAnnual; }
    public double getGuaranteedRateAnnual() { return guaranteedRateAnnual; }
}