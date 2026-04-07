import java.util.HashMap;

public class Sheet_CashValue {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();
    private final Sheet_Inputs inputs;
    private final Sheet_Rates rates;
    private final Sheet_Premium premium;

    // Public maps for Output sheet
    public final HashMap<Integer, Double> nongEndAVMap = new HashMap<>();
    public final HashMap<Integer, Double> guarEndAVMap = new HashMap<>();
    public final HashMap<Integer, Double> nongDBMap = new HashMap<>();

    public Sheet_CashValue(Sheet_Inputs inputs, Sheet_Rates rates, Sheet_Premium premium) {
        this.inputs = inputs;
        this.rates = rates;
        this.premium = premium;
    }

    public void initialize_calculation() {
        String[] columns = {"A", "B", "C", "D", "E", "G", "H", "I", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T"};
        for(String c : columns) data.put(c, new HashMap<>());

        double nongBegAV = 0.0;
        double guarBegAV = 0.0;
        
        int startRow = 2;
        int months = inputs.projectionYears * 12;

        double nongIntFactor = Math.pow(1 + inputs.currentRateAnnual, 1.0/12.0) - 1.0;
        double guarIntFactor = Math.pow(1 + inputs.guaranteedRateAnnual, 1.0/12.0) - 1.0;
        double guarCoiMultiplier = 1.5; 

        for (int row = startRow; row < startRow + months; row++) {
            
            double date = (double) premium.getCalculatedResults().get("A").get(row);
            int year = (int) premium.getCalculatedResults().get("B").get(row);
            int month = (int) premium.getCalculatedResults().get("C").get(row);
            
            int age = inputs.issueAge + (year - 1);
            
            data.get("A").put(row, date);
            data.get("B").put(row, year);
            data.get("C").put(row, month);
            data.get("D").put(row, age);
            data.get("E").put(row, age); 

            Sheet_Rates.RateRow rateRow = rates.getRates(age);
            double netPrem = premium.netPremMap.getOrDefault(row, 0.0);

            // --- NONG CALCULATION ---
            data.get("G").put(row, nongBegAV);
            data.get("H").put(row, netPrem);

            double nongBaseVal = nongBegAV + netPrem;
            double nongDB = Math.max(inputs.faceAmount, nongBaseVal * rateRow.corridorFactor);
            double nongNAR = Math.max(0, nongDB - nongBaseVal);
            
            double nongCoiCharge = nongNAR * rateRow.coiRate;
            double nongTotalCharges = inputs.monthlyAdminFee + nongCoiCharge;
            data.get("I").put(row, nongTotalCharges);

            double nongInterestBase = nongBaseVal - nongTotalCharges;
            double nongInterest = nongInterestBase * nongIntFactor;
            data.get("J").put(row, nongInterest);

            double nongEndAV = nongInterestBase + nongInterest;
            data.get("K").put(row, nongEndAV);
            nongEndAVMap.put(row, nongEndAV);

            double surrChg = inputs.faceAmount * rateRow.surrChgRate;
            data.get("L").put(row, surrChg);

            double csv = Math.max(0, nongEndAV - surrChg);
            data.get("M").put(row, csv);

            double finalNongDB = Math.max(inputs.faceAmount, nongEndAV * rateRow.corridorFactor);
            data.get("N").put(row, finalNongDB);
            nongDBMap.put(row, finalNongDB);

            // --- GUAR CALCULATION ---
            data.get("P").put(row, guarBegAV);
            data.get("Q").put(row, netPrem);

            double guarBaseVal = guarBegAV + netPrem;
            double guarDB = Math.max(inputs.faceAmount, guarBaseVal * rateRow.corridorFactor);
            double guarNAR = Math.max(0, guarDB - guarBaseVal);

            double guarCoiCharge = guarNAR * (rateRow.coiRate * guarCoiMultiplier);
            double guarTotalCharges = inputs.monthlyAdminFee + guarCoiCharge;
            data.get("R").put(row, guarTotalCharges);

            double guarInterestBase = guarBaseVal - guarTotalCharges;
            double guarInterest = guarInterestBase * guarIntFactor;
            data.get("S").put(row, guarInterest);

            double guarEndAV = guarInterestBase + guarInterest;
            data.get("T").put(row, guarEndAV);
            guarEndAVMap.put(row, guarEndAV);

            nongBegAV = nongEndAV;
            guarBegAV = guarEndAV;
        }
    }

    public HashMap<String, HashMap<Integer, Object>> getCalculatedResults() {
        return data;
    }
}