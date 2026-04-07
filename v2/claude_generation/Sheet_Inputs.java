import java.util.HashMap;

/**
 * Sheet: Inputs (INPUT)
 * All values are static constants loaded from the workbook.
 * Issue Date 2009-03-27 is stored as Excel serial (days since 1900-01-01).
 * Excel serial for 2009-03-27 = 39899
 */
public class Sheet_Inputs {

    private final HashMap<String, HashMap<Integer, Object>> data = new HashMap<>();

    public void initialize_calculation() {
        // Column A: Labels
        HashMap<Integer, Object> colA = new HashMap<>();
        colA.put(1,  "Policy Setup");
        colA.put(2,  "Face Amount");
        colA.put(3,  "Issue Age");
        colA.put(4,  "Issue Date");
        colA.put(5,  "Gender");
        colA.put(6,  "Class");
        colA.put(7,  "Projection Settings");
        colA.put(8,  "Projection Years");
        colA.put(9,  "Premiums");
        colA.put(10, "Scheduled Monthly Premium");
        colA.put(11, "Target Premium (Annual)");
        colA.put(12, "Guideline Annual Limit");
        colA.put(13, "Guideline Single Limit");
        colA.put(14, "Loads & Charges");
        colA.put(15, "Prem Load % (Up to Target)");
        colA.put(16, "Prem Load % (Excess)");
        colA.put(17, "Accumulated Cash Value");
        colA.put(18, "Monthly Admin Fee");
        colA.put(19, "Interest Rates");
        colA.put(20, "Current Rate (Annual)");
        colA.put(21, "Guaranteed Rate (Annual)");
        data.put("A", colA);

        // Column B: Values
        HashMap<Integer, Object> colB = new HashMap<>();
        // B2: Face Amount = 100000
        colB.put(2,  100000.0);
        // B3: Issue Age = 35
        colB.put(3,  35.0);
        // B4: Issue Date = 2009-03-27 → Excel serial 39899
        colB.put(4,  39899.0);
        // B5: Gender = Female
        colB.put(5,  "Female");
        // B6: Class = Non-Smoker
        colB.put(6,  "Non-Smoker");
        // B8: Projection Years = 11
        colB.put(8,  11.0);
        // B10: Scheduled Monthly Premium = 50
        colB.put(10, 50.0);
        // B11: Target Premium (Annual) = 630.71
        colB.put(11, 630.71);
        // B12: Guideline Annual Limit = 1049.05
        colB.put(12, 1049.05);
        // B13: Guideline Single Limit = 11801.53
        colB.put(13, 11801.53);
        // B15: Prem Load % (Up to Target) = 0.07
        colB.put(15, 0.07);
        // B16: Prem Load % (Excess) = 0.05
        colB.put(16, 0.05);
        // B17: Accumulated Cash Value (ACV) = 11410.54
        colB.put(17, 11410.54);
        // B18: Monthly Admin Fee = 5
        colB.put(18, 5.0);
        // B20: Current Rate (Annual) = 0.0352
        colB.put(20, 0.0352);
        // B21: Guaranteed Rate (Annual) = 0.03
        colB.put(21, 0.03);
        data.put("B", colB);
    }

    public HashMap<String, HashMap<Integer, Object>> getData() {
        return data;
    }

    // ── Convenience accessors (used by other sheets) ──────────────────────────

    /** Input_FaceAmt  → B2 */
    public double getFaceAmt()     { return (double) data.get("B").get(2); }
    /** Input_IssueAge → B3 */
    public double getIssueAge()    { return (double) data.get("B").get(3); }
    /** Input_IssueDate (Excel serial) → B4 */
    public double getIssueDate()   { return (double) data.get("B").get(4); }
    /** Input_ProjYears → B8 */
    public double getProjYears()   { return (double) data.get("B").get(8); }
    /** Input_SchedPrem → B10 */
    public double getSchedPrem()   { return (double) data.get("B").get(10); }
    /** Input_TargetPrem → B11 */
    public double getTargetPrem()  { return (double) data.get("B").get(11); }
    /** Input_Load_Tgt → B15 */
    public double getLoadTgt()     { return (double) data.get("B").get(15); }
    /** Input_Load_Exc → B16 */
    public double getLoadExc()     { return (double) data.get("B").get(16); }
    /** ACV → B17 */
    public double getACV()         { return (double) data.get("B").get(17); }
    /** Input_AdminFee → B18 */
    public double getAdminFee()    { return (double) data.get("B").get(18); }
    /** Input_Int_Cur → B20 */
    public double getIntCur()      { return (double) data.get("B").get(20); }
    /** Input_Int_Guar → B21 */
    public double getIntGuar()     { return (double) data.get("B").get(21); }

    // Derived named ranges
    /** Calc_MthlyRate_Cur = (1+Input_Int_Cur)^(1/12)-1 */
    public double getMthlyRateCur()  { return Math.pow(1.0 + getIntCur(),  1.0/12.0) - 1.0; }
    /** Calc_MthlyRate_Guar = (1+Input_Int_Guar)^(1/12)-1 */
    public double getMthlyRateGuar() { return Math.pow(1.0 + getIntGuar(), 1.0/12.0) - 1.0; }
}
