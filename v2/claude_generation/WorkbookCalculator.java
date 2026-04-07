import java.util.HashMap;

/**
 * WorkbookCalculator — Master Orchestrator
 *
 * Instantiates all sheet classes for gaulEngine_v2.xlsx in dependency order,
 * calls initialize_calculation() exactly once per sheet, and returns a
 * consolidated workbook output.
 *
 * Output structure:
 *   HashMap<sheetName, HashMap<columnLetter, HashMap<rowNumber, value>>>
 *
 * Dependency order:
 *   1. Sheet_Inputs    (no deps)
 *   2. Sheet_Rates     (no deps)
 *   3. Sheet_Premium   (depends on Inputs, Rates)
 *   4. Sheet_CashValue (depends on Inputs, Rates, Premium)
 *   5. Sheet_Output    (depends on Inputs, Premium, CashValue)
 */
public class WorkbookCalculator {

    public static HashMap<String, HashMap<String, HashMap<Integer, Object>>> calculate() {

        // 1. Inputs
        Sheet_Inputs inputs = new Sheet_Inputs();
        inputs.initialize_calculation();

        // 2. Rates
        Sheet_Rates rates = new Sheet_Rates();
        rates.initialize_calculation();

        // 3. Premium
        Sheet_Premium premium = new Sheet_Premium(inputs, rates);
        premium.initialize_calculation();

        // 4. Cash Value
        Sheet_CashValue cashValue = new Sheet_CashValue(inputs, rates, premium);
        cashValue.initialize_calculation();

        // 5. Output
        Sheet_Output output = new Sheet_Output(inputs, premium, cashValue);
        output.initialize_calculation();

        // Consolidate
        HashMap<String, HashMap<String, HashMap<Integer, Object>>> workbook = new HashMap<>();
        workbook.put("Inputs",     inputs.getData());
        workbook.put("Rates",      rates.getData());
        workbook.put("Premium",    premium.getData());
        workbook.put("Cash Value", cashValue.getOutput());
        workbook.put("Output",     output.getOutput());

        return workbook;
    }

    /**
     * Entry point for quick smoke-test / verification.
     * Prints Output sheet rows 2-12 to stdout.
     */
    public static void main(String[] args) {
        HashMap<String, HashMap<String, HashMap<Integer, Object>>> wb = calculate();

        System.out.println("=== Output Sheet ===");
        HashMap<String, HashMap<Integer, Object>> out = wb.get("Output");
        String[] cols = {"A","B","C","D","E","F","G","H"};
        String[] hdrs = {"Year","End Date","Age","Ann Prem","Cum Prem","NONG CSV","NONG DB","GUAR CSV"};

        // Header
        System.out.printf("%-6s", "Row");
        for (String h : hdrs) System.out.printf("%-18s", h);
        System.out.println();

        for (int row = 2; row <= 19; row++) {
            Object yearVal = out.get("A").get(row);
            if ("".equals(yearVal)) break;
            System.out.printf("%-6d", row);
            for (String col : cols) {
                Object v = out.get(col).get(row);
                if (v instanceof Double) {
                    System.out.printf("%-18.4f", (Double) v);
                } else {
                    System.out.printf("%-18s", v);
                }
            }
            System.out.println();
        }

        System.out.println("\n=== Premium Sheet (first 5 rows) ===");
        HashMap<String, HashMap<Integer, Object>> prem = wb.get("Premium");
        String[] pcols = {"A","B","C","D","E","F","G","H","I","J","K","L"};
        for (int row = 2; row <= 6; row++) {
            System.out.printf("Row %d: ", row);
            for (String col : pcols) {
                Object v = prem.get(col).get(row);
                if (v instanceof Double) System.out.printf("%s=%.4f  ", col, (Double)v);
                else System.out.printf("%s=%s  ", col, v);
            }
            System.out.println();
        }

        System.out.println("\n=== Cash Value Sheet (first 3 rows, NONG) ===");
        HashMap<String, HashMap<Integer, Object>> cv = wb.get("Cash Value");
        String[] cvcols = {"A","B","C","D","F","G","H","I","J","K","L","M"};
        for (int row = 2; row <= 4; row++) {
            System.out.printf("Row %d: ", row);
            for (String col : cvcols) {
                Object v = cv.get(col).get(row);
                if (v instanceof Double) System.out.printf("%s=%.4f  ", col, (Double)v);
                else System.out.printf("%s=%s  ", col, v);
            }
            System.out.println();
        }
    }
}
