import java.util.HashMap;

public class Sheet_Premium {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();
    private final Sheet_Inputs inputs;

    // Internal storage for cross-sheet access
    public final HashMap<Integer, Double> netPremMap = new HashMap<>();
    public final HashMap<Integer, Double> cumPremMap = new HashMap<>();
    public final HashMap<Integer, Double> dateMap = new HashMap<>();

    public Sheet_Premium(Sheet_Inputs inputs) {
        this.inputs = inputs;
    }

    public void initialize_calculation() {
        String[] cols = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"};
        for (String c : cols) data.put(c, new HashMap<>());

        double currentSerialDate = inputs.issueDateSerial;
        int currentPolicyYear = 1;
        int currentMonth = 1;
        double annPremYTD = 0.0;
        double cumPrem = 0.0;
        
        int startRow = 2;
        int monthsToProject = inputs.projectionYears * 12;

        for (int row = startRow; row < startRow + monthsToProject; row++) {
            
            // 1. Date (A), Year (B), Month (C)
            data.get("A").put(row, currentSerialDate);
            data.get("B").put(row, currentPolicyYear);
            data.get("C").put(row, currentMonth);
            dateMap.put(row, currentSerialDate);

            // 2. Gross Prem (D) - Constant from inputs
            double grossPrem = inputs.scheduledMonthlyPremium;
            data.get("D").put(row, grossPrem);

            // 3. Ann Prem YTD (E)
            if (currentMonth == 1) annPremYTD = 0.0;
            annPremYTD += grossPrem;
            data.get("E").put(row, annPremYTD);

            // 4. Cum Prem (F)
            cumPrem += grossPrem;
            data.get("F").put(row, cumPrem);
            cumPremMap.put(row, cumPrem);

            // 5. Target Rem (G) 
            double prevYTD = annPremYTD - grossPrem;
            double targetRem = Math.max(0, inputs.targetPremiumAnnual - prevYTD);
            data.get("G").put(row, targetRem);

            // 6. Prem Tgt (H) & Prem Exc (I)
            double premInTarget = Math.min(grossPrem, targetRem);
            double premExcess = grossPrem - premInTarget;
            data.get("H").put(row, premInTarget);
            data.get("I").put(row, premExcess);

            // 7. Load Amt (J)
            double load = (premInTarget * inputs.premLoadTarget) + (premExcess * inputs.premLoadExcess);
            data.get("J").put(row, load);

            // 8. Net Prem (K)
            double netPrem = grossPrem - load;
            data.get("K").put(row, netPrem);
            netPremMap.put(row, netPrem);

            // --- Increment Time ---
            currentSerialDate += 30.42; 
            currentMonth++;
            if (currentMonth > 12) {
                currentMonth = 1;
                currentPolicyYear++;
            }
        }
    }

    public HashMap<String, HashMap<Integer, Object>> getCalculatedResults() {
        return data;
    }
}