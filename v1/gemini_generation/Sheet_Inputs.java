import java.util.HashMap;

public class Sheet_Inputs {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();

    // Cached values for easier access by other sheets
    public final double faceAmount;
    public final int issueAge;
    public final double issueDateSerial; // Excel serial date
    public final String gender;
    public final String riskClass;
    public final int projectionYears;
    public final double scheduledMonthlyPremium;
    public final double targetPremiumAnnual;
    public final double premLoadTarget;
    public final double premLoadExcess;
    public final double monthlyAdminFee;
    public final double currentRateAnnual;
    public final double guaranteedRateAnnual;

    public Sheet_Inputs() {
        // Initialize Map Structure
        data.put("B", new HashMap<>());

        // --- Row 2: Face Amount ---
        faceAmount = 100000.0;
        putValue("B", 2, faceAmount);

        // --- Row 3: Issue Age ---
        issueAge = 35;
        putValue("B", 3, issueAge);

        // --- Row 4: Issue Date (2009-03-27) ---
        // Excel serial for 2009-03-27. 
        issueDateSerial = 39899.0; 
        putValue("B", 4, issueDateSerial); 

        // --- Row 5: Gender ---
        gender = "Female";
        putValue("B", 5, gender);

        // --- Row 6: Class ---
        riskClass = "Non-Smoker";
        putValue("B", 6, riskClass);

        // --- Row 9: Projection Years ---
        projectionYears = 5;
        putValue("B", 9, projectionYears);

        // --- Row 12: Scheduled Monthly Premium ---
        scheduledMonthlyPremium = 50.0;
        putValue("B", 12, scheduledMonthlyPremium);

        // --- Row 13: Target Premium (Annual) ---
        targetPremiumAnnual = 630.71;
        putValue("B", 13, targetPremiumAnnual);

        // --- Row 14: Guideline Annual Limit ---
        putValue("B", 14, 1049.05);

        // --- Row 15: Guideline Single Limit ---
        putValue("B", 15, 11801.53);

        // --- Row 18: Prem Load % (Up to Target) ---
        premLoadTarget = 0.07;
        putValue("B", 18, premLoadTarget);

        // --- Row 19: Prem Load % (Excess) ---
        premLoadExcess = 0.05;
        putValue("B", 19, premLoadExcess);

        // --- Row 20: Monthly Admin Fee ---
        monthlyAdminFee = 5.0;
        putValue("B", 20, monthlyAdminFee);

        // --- Row 23: Current Rate (Annual) ---
        currentRateAnnual = 0.0352;
        putValue("B", 23, currentRateAnnual);

        // --- Row 24: Guaranteed Rate (Annual) ---
        guaranteedRateAnnual = 0.03;
        putValue("B", 24, guaranteedRateAnnual);
    }

    private void putValue(String col, int row, Object val) {
        if (!data.containsKey(col)) {
            data.put(col, new HashMap<>());
        }
        data.get(col).put(row, val);
    }

    public HashMap<String, HashMap<Integer, Object>> getCalculatedResults() {
        return data;
    }
}