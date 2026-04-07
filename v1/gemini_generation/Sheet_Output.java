import java.util.HashMap;

public class Sheet_Output {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();
    private final Sheet_Premium premium;
    private final Sheet_CashValue cashValue;
    private final int projectionYears;

    public Sheet_Output(Sheet_Premium premium, Sheet_CashValue cashValue, int projectionYears) {
        this.premium = premium;
        this.cashValue = cashValue;
        this.projectionYears = projectionYears;
    }

    public void initialize_calculation() {
        String[] cols = {"A", "B", "C", "D", "E", "F", "G", "H"};
        for(String c : cols) data.put(c, new HashMap<>());

        int outputRow = 2;

        for (int year = 1; year <= projectionYears; year++) {
            
            int month12Index = 2 + ((year - 1) * 12) + 11;

            // 1. Policy Year
            data.get("A").put(outputRow, year);

            // 2. End Date
            double endDate = premium.dateMap.get(month12Index);
            data.get("B").put(outputRow, endDate);

            // 3. Attained Age
            int age = (int) cashValue.getCalculatedResults().get("D").get(month12Index);
            data.get("C").put(outputRow, age);

            // 4. Annual Premium
            double cumPremNow = premium.cumPremMap.get(month12Index);
            double cumPremPrev = (year == 1) ? 0.0 : premium.cumPremMap.get(month12Index - 12);
            data.get("D").put(outputRow, cumPremNow - cumPremPrev);

            // 5. Cumulative Premium
            data.get("E").put(outputRow, cumPremNow);

            // 6. NONG Cash Value
            double nongAv = cashValue.nongEndAVMap.get(month12Index);
            data.get("F").put(outputRow, nongAv);

            // 7. NONG Death Benefit
            double nongDb = cashValue.nongDBMap.get(month12Index);
            data.get("G").put(outputRow, nongDb);

            // 8. GUAR Cash Value
            double guarAv = cashValue.guarEndAVMap.get(month12Index);
            data.get("H").put(outputRow, guarAv);

            outputRow++;
        }
    }

    public HashMap<String, HashMap<Integer, Object>> getCalculatedResults() {
        return data;
    }
}