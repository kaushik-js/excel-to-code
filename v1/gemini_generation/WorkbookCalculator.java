import java.util.HashMap;

public class WorkbookCalculator {

    public HashMap<String, HashMap<String, HashMap<Integer, Object>>> calculateWorkbook() {
        
        Sheet_Inputs inputs = new Sheet_Inputs();
        Sheet_Rates rates = new Sheet_Rates();

        Sheet_Premium premium = new Sheet_Premium(inputs);
        Sheet_CashValue cashValue = new Sheet_CashValue(inputs, rates, premium);
        
        Sheet_Output output = new Sheet_Output(premium, cashValue, inputs.projectionYears);

        premium.initialize_calculation();
        cashValue.initialize_calculation();
        output.initialize_calculation();

        HashMap<String, HashMap<String, HashMap<Integer, Object>>> workbook = new HashMap<>();
        
        workbook.put("Inputs", inputs.getCalculatedResults());
        workbook.put("Rates", rates.getCalculatedResults());
        workbook.put("Premium", premium.getCalculatedResults());
        workbook.put("Cash Value", cashValue.getCalculatedResults());
        workbook.put("Output", output.getCalculatedResults());

        return workbook;
    }

    public static void main(String[] args) {
        WorkbookCalculator calc = new WorkbookCalculator();
        var results = calc.calculateWorkbook();
        System.out.println(results.get("Output"));
        
    }
}