import java.util.HashMap;
import java.util.Map;

public class Sheet_Rates {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();
    
    // Internal Lookup Map: Age -> RateRow
    private final HashMap<Integer, RateRow> rateTable = new HashMap<>();

    public static class RateRow {
        double coiRate;
        double surrChgRate;
        double corridorFactor;

        public RateRow(double coi, double surr, double corr) {
            this.coiRate = coi;
            this.surrChgRate = surr;
            this.corridorFactor = corr;
        }
    }

    public Sheet_Rates() {
        loadRates();
        
        data.put("A", new HashMap<>()); // Age
        data.put("B", new HashMap<>()); // COI
        data.put("C", new HashMap<>()); // Surr
        data.put("D", new HashMap<>()); // Corridor

        for (Map.Entry<Integer, RateRow> entry : rateTable.entrySet()) {
            int row = entry.getKey() + 2; 
            data.get("A").put(row, entry.getKey());
            data.get("B").put(row, entry.getValue().coiRate);
            data.get("C").put(row, entry.getValue().surrChgRate);
            data.get("D").put(row, entry.getValue().corridorFactor);
        }
    }

    public RateRow getRates(int age) {
        return rateTable.getOrDefault(age, new RateRow(0.0, 0.0, 1.0));
    }

    private void loadRates() {
        // Data derived from the "Rates.csv" provided
        addRate(35, 0.00013593382300394741, 0.05, 2.5);
        addRate(36, 0.0001447695214992039, 0.05, 2.5);
        addRate(37, 0.00015417954039665221, 0.05, 2.5);
        addRate(38, 0.0001642012105224345, 0.05, 2.5);
        addRate(39, 0.0001748742892063928, 0.05, 2.5);
        addRate(40, 0.00018624111800480831, 0.05, 2.5);
        addRate(41, 0.0001983467906751209, 0.05, 2.4727272727272731);
        
        // Dummy logic for unlisted ages
        for(int i=0; i<35; i++) addRate(i, 0.00001, 0.05, 2.5);
        for(int i=42; i<100; i++) addRate(i, 0.0002 + (i*0.00001), 0.05, 1.0);
    }

    private void addRate(int age, double coi, double surr, double corr) {
        rateTable.put(age, new RateRow(coi, surr, corr));
    }

    public HashMap<String, HashMap<Integer, Object>> getCalculatedResults() {
        return data;
    }
}