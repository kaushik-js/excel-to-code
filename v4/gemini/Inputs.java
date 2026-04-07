public class Inputs {
    // Policy Setup
    public final double faceAmount = 100000.0;
    public final int issueAge = 35;
    public final String gender = "Female";
    public final String riskClass = "Non-Smoker";
    
    // Projection Settings
    public final int projectionYears = 11;
    
    // Premiums
    public final double scheduledMonthlyPremium = 50.0;
    public final double targetPremiumAnnual = 630.71;
    public final double guidelineAnnualLimit = 1049.05;
    public final double guidelineSingleLimit = 11801.53;
    
    // Loads & Charges
    public final double premLoadUpToTarget = 0.07;
    public final double premLoadExcess = 0.05;
    public final double accumulatedCashValue = 11410.54; // Initial Beg AV
    public final double monthlyAdminFee = 5.0;
    
    // Interest Rates
    public final double currentRateAnnual = 0.0352;
    public final double guaranteedRateAnnual = 0.03;
}