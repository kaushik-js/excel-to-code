public class CashValueSheet {

    public double cell_A2() {
        return Input_IssueDate;
    }

    public double cell_B2() {
        return 1;
    }

    public double cell_C2() {
        return 1;
    }

    public double cell_D2() {
        return Input_IssueAge+(B2-1);
    }

    public double cell_F2() {
        return ACV;
    }

    public double cell_G2() {
        return VLOOKUP(A2, Premium!1:1048, 12, TRUE);
    }

    public double cell_H2() {
        return ExcelFunctions.IF(B2="", 0, (VLOOKUP(D2, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F2+G2))) + Input_AdminFee);
    }

    public double cell_I2() {
        return ExcelFunctions.IF(B2="", 0, (F2+G2-H2) * Calc_MthlyRate_Cur);
    }

    public double cell_J2() {
        return ExcelFunctions.IF(B2="", 0, Math.max(0, F2+G2-H2+I2));
    }

    public double cell_K2() {
        return ExcelFunctions.IF(B2="", 0, VLOOKUP(D2, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L2() {
        return ExcelFunctions.IF(B2="", 0, Math.max(0, J2-K2));
    }

    public double cell_M2() {
        return ExcelFunctions.IF(B2="", 0, Math.max(Input_FaceAmt, J2 * VLOOKUP(D2, Table_Rates, 4, FALSE)));
    }

    public double cell_O2() {
        return ACV;
    }

    public double cell_P2() {
        return Premium!L2;
    }

    public double cell_Q2() {
        return ExcelFunctions.IF(B2="", 0, ((VLOOKUP(D2, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O2+P2))) + Input_AdminFee);
    }

    public double cell_R2() {
        return ExcelFunctions.IF(B2="", 0, (O2+P2-Q2) * Calc_MthlyRate_Guar);
    }

    public double cell_S2() {
        return ExcelFunctions.IF(B2="", 0, Math.max(0, O2+P2-Q2+R2));
    }

    public double cell_A3() {
        return EDATE(A2, 1);
    }

    public double cell_B3() {
        return ExcelFunctions.IF(B2>Input_ProjYears, "", ExcelFunctions.IF(C2=12, B2+1, B2));
    }

    public double cell_C3() {
        return ExcelFunctions.IF(C2=12, 1, C2+1);
    }

    public double cell_D3() {
        return Input_IssueAge+(B3-1);
    }

    public double cell_F3() {
        return ExcelFunctions.IF(B3="", 0, J2);
    }

    public double cell_G3() {
        return VLOOKUP(A3, Premium!2:1049, 12, TRUE);
    }

    public double cell_H3() {
        return ExcelFunctions.IF(B3="", 0, (VLOOKUP(D3, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F3+G3))) + Input_AdminFee);
    }

    public double cell_I3() {
        return ExcelFunctions.IF(B3="", 0, (F3+G3-H3) * Calc_MthlyRate_Cur);
    }

    public double cell_J3() {
        return ExcelFunctions.IF(B3="", 0, Math.max(0, F3+G3-H3+I3));
    }

    public double cell_K3() {
        return ExcelFunctions.IF(B3="", 0, VLOOKUP(D3, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L3() {
        return ExcelFunctions.IF(B3="", 0, Math.max(0, J3-K3));
    }

    public double cell_M3() {
        return ExcelFunctions.IF(B3="", 0, Math.max(Input_FaceAmt, J3 * VLOOKUP(D3, Table_Rates, 4, FALSE)));
    }

    public double cell_O3() {
        return ACV;
    }

    public double cell_P3() {
        return Premium!L3;
    }

    public double cell_Q3() {
        return ExcelFunctions.IF(B3="", 0, ((VLOOKUP(D3, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O3+P3))) + Input_AdminFee);
    }

    public double cell_R3() {
        return ExcelFunctions.IF(B3="", 0, (O3+P3-Q3) * Calc_MthlyRate_Guar);
    }

    public double cell_S3() {
        return ExcelFunctions.IF(B3="", 0, Math.max(0, O3+P3-Q3+R3));
    }

    public double cell_A4() {
        return EDATE(A3, 1);
    }

    public double cell_B4() {
        return ExcelFunctions.IF(B3>=Input_ProjYears, "", ExcelFunctions.IF(C3=12, B3+1, B3));
    }

    public double cell_C4() {
        return ExcelFunctions.IF(C3=12, 1, C3+1);
    }

    public double cell_D4() {
        return Input_IssueAge+(B4-1);
    }

    public double cell_F4() {
        return ExcelFunctions.IF(B4="", 0, J3);
    }

    public double cell_G4() {
        return VLOOKUP(A4, Premium!3:1050, 12, TRUE);
    }

    public double cell_H4() {
        return ExcelFunctions.IF(B4="", 0, (VLOOKUP(D4, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F4+G4))) + Input_AdminFee);
    }

    public double cell_I4() {
        return ExcelFunctions.IF(B4="", 0, (F4+G4-H4) * Calc_MthlyRate_Cur);
    }

    public double cell_J4() {
        return ExcelFunctions.IF(B4="", 0, Math.max(0, F4+G4-H4+I4));
    }

    public double cell_K4() {
        return ExcelFunctions.IF(B4="", 0, VLOOKUP(D4, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L4() {
        return ExcelFunctions.IF(B4="", 0, Math.max(0, J4-K4));
    }

    public double cell_M4() {
        return ExcelFunctions.IF(B4="", 0, Math.max(Input_FaceAmt, J4 * VLOOKUP(D4, Table_Rates, 4, FALSE)));
    }

    public double cell_O4() {
        return ACV;
    }

    public double cell_P4() {
        return Premium!L4;
    }

    public double cell_Q4() {
        return ExcelFunctions.IF(B4="", 0, ((VLOOKUP(D4, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O4+P4))) + Input_AdminFee);
    }

    public double cell_R4() {
        return ExcelFunctions.IF(B4="", 0, (O4+P4-Q4) * Calc_MthlyRate_Guar);
    }

    public double cell_S4() {
        return ExcelFunctions.IF(B4="", 0, Math.max(0, O4+P4-Q4+R4));
    }

    public double cell_A5() {
        return EDATE(A4, 1);
    }

    public double cell_B5() {
        return ExcelFunctions.IF(B4>=Input_ProjYears, "", ExcelFunctions.IF(C4=12, B4+1, B4));
    }

    public double cell_C5() {
        return ExcelFunctions.IF(C4=12, 1, C4+1);
    }

    public double cell_D5() {
        return Input_IssueAge+(B5-1);
    }

    public double cell_F5() {
        return ExcelFunctions.IF(B5="", 0, J4);
    }

    public double cell_G5() {
        return VLOOKUP(A5, Premium!4:1051, 12, TRUE);
    }

    public double cell_H5() {
        return ExcelFunctions.IF(B5="", 0, (VLOOKUP(D5, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F5+G5))) + Input_AdminFee);
    }

    public double cell_I5() {
        return ExcelFunctions.IF(B5="", 0, (F5+G5-H5) * Calc_MthlyRate_Cur);
    }

    public double cell_J5() {
        return ExcelFunctions.IF(B5="", 0, Math.max(0, F5+G5-H5+I5));
    }

    public double cell_K5() {
        return ExcelFunctions.IF(B5="", 0, VLOOKUP(D5, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L5() {
        return ExcelFunctions.IF(B5="", 0, Math.max(0, J5-K5));
    }

    public double cell_M5() {
        return ExcelFunctions.IF(B5="", 0, Math.max(Input_FaceAmt, J5 * VLOOKUP(D5, Table_Rates, 4, FALSE)));
    }

    public double cell_O5() {
        return ACV;
    }

    public double cell_P5() {
        return Premium!L5;
    }

    public double cell_Q5() {
        return ExcelFunctions.IF(B5="", 0, ((VLOOKUP(D5, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O5+P5))) + Input_AdminFee);
    }

    public double cell_R5() {
        return ExcelFunctions.IF(B5="", 0, (O5+P5-Q5) * Calc_MthlyRate_Guar);
    }

    public double cell_S5() {
        return ExcelFunctions.IF(B5="", 0, Math.max(0, O5+P5-Q5+R5));
    }

    public double cell_A6() {
        return EDATE(A5, 1);
    }

    public double cell_B6() {
        return ExcelFunctions.IF(B5>=Input_ProjYears, "", ExcelFunctions.IF(C5=12, B5+1, B5));
    }

    public double cell_C6() {
        return ExcelFunctions.IF(C5=12, 1, C5+1);
    }

    public double cell_D6() {
        return Input_IssueAge+(B6-1);
    }

    public double cell_F6() {
        return ExcelFunctions.IF(B6="", 0, J5);
    }

    public double cell_G6() {
        return VLOOKUP(A6, Premium!5:1052, 12, TRUE);
    }

    public double cell_H6() {
        return ExcelFunctions.IF(B6="", 0, (VLOOKUP(D6, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F6+G6))) + Input_AdminFee);
    }

    public double cell_I6() {
        return ExcelFunctions.IF(B6="", 0, (F6+G6-H6) * Calc_MthlyRate_Cur);
    }

    public double cell_J6() {
        return ExcelFunctions.IF(B6="", 0, Math.max(0, F6+G6-H6+I6));
    }

    public double cell_K6() {
        return ExcelFunctions.IF(B6="", 0, VLOOKUP(D6, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L6() {
        return ExcelFunctions.IF(B6="", 0, Math.max(0, J6-K6));
    }

    public double cell_M6() {
        return ExcelFunctions.IF(B6="", 0, Math.max(Input_FaceAmt, J6 * VLOOKUP(D6, Table_Rates, 4, FALSE)));
    }

    public double cell_O6() {
        return ACV;
    }

    public double cell_P6() {
        return Premium!L6;
    }

    public double cell_Q6() {
        return ExcelFunctions.IF(B6="", 0, ((VLOOKUP(D6, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O6+P6))) + Input_AdminFee);
    }

    public double cell_R6() {
        return ExcelFunctions.IF(B6="", 0, (O6+P6-Q6) * Calc_MthlyRate_Guar);
    }

    public double cell_S6() {
        return ExcelFunctions.IF(B6="", 0, Math.max(0, O6+P6-Q6+R6));
    }

    public double cell_A7() {
        return EDATE(A6, 1);
    }

    public double cell_B7() {
        return ExcelFunctions.IF(B6>=Input_ProjYears, "", ExcelFunctions.IF(C6=12, B6+1, B6));
    }

    public double cell_C7() {
        return ExcelFunctions.IF(C6=12, 1, C6+1);
    }

    public double cell_D7() {
        return Input_IssueAge+(B7-1);
    }

    public double cell_F7() {
        return ExcelFunctions.IF(B7="", 0, J6);
    }

    public double cell_G7() {
        return VLOOKUP(A7, Premium!6:1053, 12, TRUE);
    }

    public double cell_H7() {
        return ExcelFunctions.IF(B7="", 0, (VLOOKUP(D7, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F7+G7))) + Input_AdminFee);
    }

    public double cell_I7() {
        return ExcelFunctions.IF(B7="", 0, (F7+G7-H7) * Calc_MthlyRate_Cur);
    }

    public double cell_J7() {
        return ExcelFunctions.IF(B7="", 0, Math.max(0, F7+G7-H7+I7));
    }

    public double cell_K7() {
        return ExcelFunctions.IF(B7="", 0, VLOOKUP(D7, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L7() {
        return ExcelFunctions.IF(B7="", 0, Math.max(0, J7-K7));
    }

    public double cell_M7() {
        return ExcelFunctions.IF(B7="", 0, Math.max(Input_FaceAmt, J7 * VLOOKUP(D7, Table_Rates, 4, FALSE)));
    }

    public double cell_O7() {
        return ACV;
    }

    public double cell_P7() {
        return Premium!L7;
    }

    public double cell_Q7() {
        return ExcelFunctions.IF(B7="", 0, ((VLOOKUP(D7, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O7+P7))) + Input_AdminFee);
    }

    public double cell_R7() {
        return ExcelFunctions.IF(B7="", 0, (O7+P7-Q7) * Calc_MthlyRate_Guar);
    }

    public double cell_S7() {
        return ExcelFunctions.IF(B7="", 0, Math.max(0, O7+P7-Q7+R7));
    }

    public double cell_A8() {
        return EDATE(A7, 1);
    }

    public double cell_B8() {
        return ExcelFunctions.IF(B7>=Input_ProjYears, "", ExcelFunctions.IF(C7=12, B7+1, B7));
    }

    public double cell_C8() {
        return ExcelFunctions.IF(C7=12, 1, C7+1);
    }

    public double cell_D8() {
        return Input_IssueAge+(B8-1);
    }

    public double cell_F8() {
        return ExcelFunctions.IF(B8="", 0, J7);
    }

    public double cell_G8() {
        return VLOOKUP(A8, Premium!7:1054, 12, TRUE);
    }

    public double cell_H8() {
        return ExcelFunctions.IF(B8="", 0, (VLOOKUP(D8, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F8+G8))) + Input_AdminFee);
    }

    public double cell_I8() {
        return ExcelFunctions.IF(B8="", 0, (F8+G8-H8) * Calc_MthlyRate_Cur);
    }

    public double cell_J8() {
        return ExcelFunctions.IF(B8="", 0, Math.max(0, F8+G8-H8+I8));
    }

    public double cell_K8() {
        return ExcelFunctions.IF(B8="", 0, VLOOKUP(D8, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L8() {
        return ExcelFunctions.IF(B8="", 0, Math.max(0, J8-K8));
    }

    public double cell_M8() {
        return ExcelFunctions.IF(B8="", 0, Math.max(Input_FaceAmt, J8 * VLOOKUP(D8, Table_Rates, 4, FALSE)));
    }

    public double cell_O8() {
        return ACV;
    }

    public double cell_P8() {
        return Premium!L8;
    }

    public double cell_Q8() {
        return ExcelFunctions.IF(B8="", 0, ((VLOOKUP(D8, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O8+P8))) + Input_AdminFee);
    }

    public double cell_R8() {
        return ExcelFunctions.IF(B8="", 0, (O8+P8-Q8) * Calc_MthlyRate_Guar);
    }

    public double cell_S8() {
        return ExcelFunctions.IF(B8="", 0, Math.max(0, O8+P8-Q8+R8));
    }

    public double cell_A9() {
        return EDATE(A8, 1);
    }

    public double cell_B9() {
        return ExcelFunctions.IF(B8>=Input_ProjYears, "", ExcelFunctions.IF(C8=12, B8+1, B8));
    }

    public double cell_C9() {
        return ExcelFunctions.IF(C8=12, 1, C8+1);
    }

    public double cell_D9() {
        return Input_IssueAge+(B9-1);
    }

    public double cell_F9() {
        return ExcelFunctions.IF(B9="", 0, J8);
    }

    public double cell_G9() {
        return VLOOKUP(A9, Premium!8:1055, 12, TRUE);
    }

    public double cell_H9() {
        return ExcelFunctions.IF(B9="", 0, (VLOOKUP(D9, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F9+G9))) + Input_AdminFee);
    }

    public double cell_I9() {
        return ExcelFunctions.IF(B9="", 0, (F9+G9-H9) * Calc_MthlyRate_Cur);
    }

    public double cell_J9() {
        return ExcelFunctions.IF(B9="", 0, Math.max(0, F9+G9-H9+I9));
    }

    public double cell_K9() {
        return ExcelFunctions.IF(B9="", 0, VLOOKUP(D9, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L9() {
        return ExcelFunctions.IF(B9="", 0, Math.max(0, J9-K9));
    }

    public double cell_M9() {
        return ExcelFunctions.IF(B9="", 0, Math.max(Input_FaceAmt, J9 * VLOOKUP(D9, Table_Rates, 4, FALSE)));
    }

    public double cell_O9() {
        return ACV;
    }

    public double cell_P9() {
        return Premium!L9;
    }

    public double cell_Q9() {
        return ExcelFunctions.IF(B9="", 0, ((VLOOKUP(D9, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O9+P9))) + Input_AdminFee);
    }

    public double cell_R9() {
        return ExcelFunctions.IF(B9="", 0, (O9+P9-Q9) * Calc_MthlyRate_Guar);
    }

    public double cell_S9() {
        return ExcelFunctions.IF(B9="", 0, Math.max(0, O9+P9-Q9+R9));
    }

    public double cell_A10() {
        return EDATE(A9, 1);
    }

    public double cell_B10() {
        return ExcelFunctions.IF(B9>=Input_ProjYears, "", ExcelFunctions.IF(C9=12, B9+1, B9));
    }

    public double cell_C10() {
        return ExcelFunctions.IF(C9=12, 1, C9+1);
    }

    public double cell_D10() {
        return Input_IssueAge+(B10-1);
    }

    public double cell_F10() {
        return ExcelFunctions.IF(B10="", 0, J9);
    }

    public double cell_G10() {
        return VLOOKUP(A10, Premium!9:1056, 12, TRUE);
    }

    public double cell_H10() {
        return ExcelFunctions.IF(B10="", 0, (VLOOKUP(D10, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F10+G10))) + Input_AdminFee);
    }

    public double cell_I10() {
        return ExcelFunctions.IF(B10="", 0, (F10+G10-H10) * Calc_MthlyRate_Cur);
    }

    public double cell_J10() {
        return ExcelFunctions.IF(B10="", 0, Math.max(0, F10+G10-H10+I10));
    }

    public double cell_K10() {
        return ExcelFunctions.IF(B10="", 0, VLOOKUP(D10, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L10() {
        return ExcelFunctions.IF(B10="", 0, Math.max(0, J10-K10));
    }

    public double cell_M10() {
        return ExcelFunctions.IF(B10="", 0, Math.max(Input_FaceAmt, J10 * VLOOKUP(D10, Table_Rates, 4, FALSE)));
    }

    public double cell_O10() {
        return ACV;
    }

    public double cell_P10() {
        return Premium!L10;
    }

    public double cell_Q10() {
        return ExcelFunctions.IF(B10="", 0, ((VLOOKUP(D10, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O10+P10))) + Input_AdminFee);
    }

    public double cell_R10() {
        return ExcelFunctions.IF(B10="", 0, (O10+P10-Q10) * Calc_MthlyRate_Guar);
    }

    public double cell_S10() {
        return ExcelFunctions.IF(B10="", 0, Math.max(0, O10+P10-Q10+R10));
    }

    public double cell_A11() {
        return EDATE(A10, 1);
    }

    public double cell_B11() {
        return ExcelFunctions.IF(B10>=Input_ProjYears, "", ExcelFunctions.IF(C10=12, B10+1, B10));
    }

    public double cell_C11() {
        return ExcelFunctions.IF(C10=12, 1, C10+1);
    }

    public double cell_D11() {
        return Input_IssueAge+(B11-1);
    }

    public double cell_F11() {
        return ExcelFunctions.IF(B11="", 0, J10);
    }

    public double cell_G11() {
        return VLOOKUP(A11, Premium!10:1057, 12, TRUE);
    }

    public double cell_H11() {
        return ExcelFunctions.IF(B11="", 0, (VLOOKUP(D11, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F11+G11))) + Input_AdminFee);
    }

    public double cell_I11() {
        return ExcelFunctions.IF(B11="", 0, (F11+G11-H11) * Calc_MthlyRate_Cur);
    }

    public double cell_J11() {
        return ExcelFunctions.IF(B11="", 0, Math.max(0, F11+G11-H11+I11));
    }

    public double cell_K11() {
        return ExcelFunctions.IF(B11="", 0, VLOOKUP(D11, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L11() {
        return ExcelFunctions.IF(B11="", 0, Math.max(0, J11-K11));
    }

    public double cell_M11() {
        return ExcelFunctions.IF(B11="", 0, Math.max(Input_FaceAmt, J11 * VLOOKUP(D11, Table_Rates, 4, FALSE)));
    }

    public double cell_O11() {
        return ACV;
    }

    public double cell_P11() {
        return Premium!L11;
    }

    public double cell_Q11() {
        return ExcelFunctions.IF(B11="", 0, ((VLOOKUP(D11, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O11+P11))) + Input_AdminFee);
    }

    public double cell_R11() {
        return ExcelFunctions.IF(B11="", 0, (O11+P11-Q11) * Calc_MthlyRate_Guar);
    }

    public double cell_S11() {
        return ExcelFunctions.IF(B11="", 0, Math.max(0, O11+P11-Q11+R11));
    }

    public double cell_A12() {
        return EDATE(A11, 1);
    }

    public double cell_B12() {
        return ExcelFunctions.IF(B11>=Input_ProjYears, "", ExcelFunctions.IF(C11=12, B11+1, B11));
    }

    public double cell_C12() {
        return ExcelFunctions.IF(C11=12, 1, C11+1);
    }

    public double cell_D12() {
        return Input_IssueAge+(B12-1);
    }

    public double cell_F12() {
        return ExcelFunctions.IF(B12="", 0, J11);
    }

    public double cell_G12() {
        return VLOOKUP(A12, Premium!11:1058, 12, TRUE);
    }

    public double cell_H12() {
        return ExcelFunctions.IF(B12="", 0, (VLOOKUP(D12, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F12+G12))) + Input_AdminFee);
    }

    public double cell_I12() {
        return ExcelFunctions.IF(B12="", 0, (F12+G12-H12) * Calc_MthlyRate_Cur);
    }

    public double cell_J12() {
        return ExcelFunctions.IF(B12="", 0, Math.max(0, F12+G12-H12+I12));
    }

    public double cell_K12() {
        return ExcelFunctions.IF(B12="", 0, VLOOKUP(D12, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L12() {
        return ExcelFunctions.IF(B12="", 0, Math.max(0, J12-K12));
    }

    public double cell_M12() {
        return ExcelFunctions.IF(B12="", 0, Math.max(Input_FaceAmt, J12 * VLOOKUP(D12, Table_Rates, 4, FALSE)));
    }

    public double cell_O12() {
        return ACV;
    }

    public double cell_P12() {
        return Premium!L12;
    }

    public double cell_Q12() {
        return ExcelFunctions.IF(B12="", 0, ((VLOOKUP(D12, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O12+P12))) + Input_AdminFee);
    }

    public double cell_R12() {
        return ExcelFunctions.IF(B12="", 0, (O12+P12-Q12) * Calc_MthlyRate_Guar);
    }

    public double cell_S12() {
        return ExcelFunctions.IF(B12="", 0, Math.max(0, O12+P12-Q12+R12));
    }

    public double cell_A13() {
        return EDATE(A12, 1);
    }

    public double cell_B13() {
        return ExcelFunctions.IF(B12>=Input_ProjYears, "", ExcelFunctions.IF(C12=12, B12+1, B12));
    }

    public double cell_C13() {
        return ExcelFunctions.IF(C12=12, 1, C12+1);
    }

    public double cell_D13() {
        return Input_IssueAge+(B13-1);
    }

    public double cell_F13() {
        return ExcelFunctions.IF(B13="", 0, J12);
    }

    public double cell_G13() {
        return VLOOKUP(A13, Premium!12:1059, 12, TRUE);
    }

    public double cell_H13() {
        return ExcelFunctions.IF(B13="", 0, (VLOOKUP(D13, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F13+G13))) + Input_AdminFee);
    }

    public double cell_I13() {
        return ExcelFunctions.IF(B13="", 0, (F13+G13-H13) * Calc_MthlyRate_Cur);
    }

    public double cell_J13() {
        return ExcelFunctions.IF(B13="", 0, Math.max(0, F13+G13-H13+I13));
    }

    public double cell_K13() {
        return ExcelFunctions.IF(B13="", 0, VLOOKUP(D13, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L13() {
        return ExcelFunctions.IF(B13="", 0, Math.max(0, J13-K13));
    }

    public double cell_M13() {
        return ExcelFunctions.IF(B13="", 0, Math.max(Input_FaceAmt, J13 * VLOOKUP(D13, Table_Rates, 4, FALSE)));
    }

    public double cell_O13() {
        return ACV;
    }

    public double cell_P13() {
        return Premium!L13;
    }

    public double cell_Q13() {
        return ExcelFunctions.IF(B13="", 0, ((VLOOKUP(D13, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O13+P13))) + Input_AdminFee);
    }

    public double cell_R13() {
        return ExcelFunctions.IF(B13="", 0, (O13+P13-Q13) * Calc_MthlyRate_Guar);
    }

    public double cell_S13() {
        return ExcelFunctions.IF(B13="", 0, Math.max(0, O13+P13-Q13+R13));
    }

    public double cell_A14() {
        return EDATE(A13, 1);
    }

    public double cell_B14() {
        return ExcelFunctions.IF(B13>=Input_ProjYears, "", ExcelFunctions.IF(C13=12, B13+1, B13));
    }

    public double cell_C14() {
        return ExcelFunctions.IF(C13=12, 1, C13+1);
    }

    public double cell_D14() {
        return Input_IssueAge+(B14-1);
    }

    public double cell_F14() {
        return ExcelFunctions.IF(B14="", 0, J13);
    }

    public double cell_G14() {
        return VLOOKUP(A14, Premium!13:1060, 12, TRUE);
    }

    public double cell_H14() {
        return ExcelFunctions.IF(B14="", 0, (VLOOKUP(D14, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F14+G14))) + Input_AdminFee);
    }

    public double cell_I14() {
        return ExcelFunctions.IF(B14="", 0, (F14+G14-H14) * Calc_MthlyRate_Cur);
    }

    public double cell_J14() {
        return ExcelFunctions.IF(B14="", 0, Math.max(0, F14+G14-H14+I14));
    }

    public double cell_K14() {
        return ExcelFunctions.IF(B14="", 0, VLOOKUP(D14, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L14() {
        return ExcelFunctions.IF(B14="", 0, Math.max(0, J14-K14));
    }

    public double cell_M14() {
        return ExcelFunctions.IF(B14="", 0, Math.max(Input_FaceAmt, J14 * VLOOKUP(D14, Table_Rates, 4, FALSE)));
    }

    public double cell_O14() {
        return ACV;
    }

    public double cell_P14() {
        return Premium!L14;
    }

    public double cell_Q14() {
        return ExcelFunctions.IF(B14="", 0, ((VLOOKUP(D14, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O14+P14))) + Input_AdminFee);
    }

    public double cell_R14() {
        return ExcelFunctions.IF(B14="", 0, (O14+P14-Q14) * Calc_MthlyRate_Guar);
    }

    public double cell_S14() {
        return ExcelFunctions.IF(B14="", 0, Math.max(0, O14+P14-Q14+R14));
    }

    public double cell_A15() {
        return EDATE(A14, 1);
    }

    public double cell_B15() {
        return ExcelFunctions.IF(B14>=Input_ProjYears, "", ExcelFunctions.IF(C14=12, B14+1, B14));
    }

    public double cell_C15() {
        return ExcelFunctions.IF(C14=12, 1, C14+1);
    }

    public double cell_D15() {
        return Input_IssueAge+(B15-1);
    }

    public double cell_F15() {
        return ExcelFunctions.IF(B15="", 0, J14);
    }

    public double cell_G15() {
        return VLOOKUP(A15, Premium!14:1061, 12, TRUE);
    }

    public double cell_H15() {
        return ExcelFunctions.IF(B15="", 0, (VLOOKUP(D15, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F15+G15))) + Input_AdminFee);
    }

    public double cell_I15() {
        return ExcelFunctions.IF(B15="", 0, (F15+G15-H15) * Calc_MthlyRate_Cur);
    }

    public double cell_J15() {
        return ExcelFunctions.IF(B15="", 0, Math.max(0, F15+G15-H15+I15));
    }

    public double cell_K15() {
        return ExcelFunctions.IF(B15="", 0, VLOOKUP(D15, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L15() {
        return ExcelFunctions.IF(B15="", 0, Math.max(0, J15-K15));
    }

    public double cell_M15() {
        return ExcelFunctions.IF(B15="", 0, Math.max(Input_FaceAmt, J15 * VLOOKUP(D15, Table_Rates, 4, FALSE)));
    }

    public double cell_O15() {
        return ACV;
    }

    public double cell_P15() {
        return Premium!L15;
    }

    public double cell_Q15() {
        return ExcelFunctions.IF(B15="", 0, ((VLOOKUP(D15, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O15+P15))) + Input_AdminFee);
    }

    public double cell_R15() {
        return ExcelFunctions.IF(B15="", 0, (O15+P15-Q15) * Calc_MthlyRate_Guar);
    }

    public double cell_S15() {
        return ExcelFunctions.IF(B15="", 0, Math.max(0, O15+P15-Q15+R15));
    }

    public double cell_A16() {
        return EDATE(A15, 1);
    }

    public double cell_B16() {
        return ExcelFunctions.IF(B15>=Input_ProjYears, "", ExcelFunctions.IF(C15=12, B15+1, B15));
    }

    public double cell_C16() {
        return ExcelFunctions.IF(C15=12, 1, C15+1);
    }

    public double cell_D16() {
        return Input_IssueAge+(B16-1);
    }

    public double cell_F16() {
        return ExcelFunctions.IF(B16="", 0, J15);
    }

    public double cell_G16() {
        return VLOOKUP(A16, Premium!15:1062, 12, TRUE);
    }

    public double cell_H16() {
        return ExcelFunctions.IF(B16="", 0, (VLOOKUP(D16, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F16+G16))) + Input_AdminFee);
    }

    public double cell_I16() {
        return ExcelFunctions.IF(B16="", 0, (F16+G16-H16) * Calc_MthlyRate_Cur);
    }

    public double cell_J16() {
        return ExcelFunctions.IF(B16="", 0, Math.max(0, F16+G16-H16+I16));
    }

    public double cell_K16() {
        return ExcelFunctions.IF(B16="", 0, VLOOKUP(D16, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L16() {
        return ExcelFunctions.IF(B16="", 0, Math.max(0, J16-K16));
    }

    public double cell_M16() {
        return ExcelFunctions.IF(B16="", 0, Math.max(Input_FaceAmt, J16 * VLOOKUP(D16, Table_Rates, 4, FALSE)));
    }

    public double cell_O16() {
        return ACV;
    }

    public double cell_P16() {
        return Premium!L16;
    }

    public double cell_Q16() {
        return ExcelFunctions.IF(B16="", 0, ((VLOOKUP(D16, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O16+P16))) + Input_AdminFee);
    }

    public double cell_R16() {
        return ExcelFunctions.IF(B16="", 0, (O16+P16-Q16) * Calc_MthlyRate_Guar);
    }

    public double cell_S16() {
        return ExcelFunctions.IF(B16="", 0, Math.max(0, O16+P16-Q16+R16));
    }

    public double cell_A17() {
        return EDATE(A16, 1);
    }

    public double cell_B17() {
        return ExcelFunctions.IF(B16>=Input_ProjYears, "", ExcelFunctions.IF(C16=12, B16+1, B16));
    }

    public double cell_C17() {
        return ExcelFunctions.IF(C16=12, 1, C16+1);
    }

    public double cell_D17() {
        return Input_IssueAge+(B17-1);
    }

    public double cell_F17() {
        return ExcelFunctions.IF(B17="", 0, J16);
    }

    public double cell_G17() {
        return VLOOKUP(A17, Premium!16:1063, 12, TRUE);
    }

    public double cell_H17() {
        return ExcelFunctions.IF(B17="", 0, (VLOOKUP(D17, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F17+G17))) + Input_AdminFee);
    }

    public double cell_I17() {
        return ExcelFunctions.IF(B17="", 0, (F17+G17-H17) * Calc_MthlyRate_Cur);
    }

    public double cell_J17() {
        return ExcelFunctions.IF(B17="", 0, Math.max(0, F17+G17-H17+I17));
    }

    public double cell_K17() {
        return ExcelFunctions.IF(B17="", 0, VLOOKUP(D17, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L17() {
        return ExcelFunctions.IF(B17="", 0, Math.max(0, J17-K17));
    }

    public double cell_M17() {
        return ExcelFunctions.IF(B17="", 0, Math.max(Input_FaceAmt, J17 * VLOOKUP(D17, Table_Rates, 4, FALSE)));
    }

    public double cell_O17() {
        return ACV;
    }

    public double cell_P17() {
        return Premium!L17;
    }

    public double cell_Q17() {
        return ExcelFunctions.IF(B17="", 0, ((VLOOKUP(D17, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O17+P17))) + Input_AdminFee);
    }

    public double cell_R17() {
        return ExcelFunctions.IF(B17="", 0, (O17+P17-Q17) * Calc_MthlyRate_Guar);
    }

    public double cell_S17() {
        return ExcelFunctions.IF(B17="", 0, Math.max(0, O17+P17-Q17+R17));
    }

    public double cell_A18() {
        return EDATE(A17, 1);
    }

    public double cell_B18() {
        return ExcelFunctions.IF(B17>=Input_ProjYears, "", ExcelFunctions.IF(C17=12, B17+1, B17));
    }

    public double cell_C18() {
        return ExcelFunctions.IF(C17=12, 1, C17+1);
    }

    public double cell_D18() {
        return Input_IssueAge+(B18-1);
    }

    public double cell_F18() {
        return ExcelFunctions.IF(B18="", 0, J17);
    }

    public double cell_G18() {
        return VLOOKUP(A18, Premium!17:1064, 12, TRUE);
    }

    public double cell_H18() {
        return ExcelFunctions.IF(B18="", 0, (VLOOKUP(D18, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F18+G18))) + Input_AdminFee);
    }

    public double cell_I18() {
        return ExcelFunctions.IF(B18="", 0, (F18+G18-H18) * Calc_MthlyRate_Cur);
    }

    public double cell_J18() {
        return ExcelFunctions.IF(B18="", 0, Math.max(0, F18+G18-H18+I18));
    }

    public double cell_K18() {
        return ExcelFunctions.IF(B18="", 0, VLOOKUP(D18, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L18() {
        return ExcelFunctions.IF(B18="", 0, Math.max(0, J18-K18));
    }

    public double cell_M18() {
        return ExcelFunctions.IF(B18="", 0, Math.max(Input_FaceAmt, J18 * VLOOKUP(D18, Table_Rates, 4, FALSE)));
    }

    public double cell_O18() {
        return ACV;
    }

    public double cell_P18() {
        return Premium!L18;
    }

    public double cell_Q18() {
        return ExcelFunctions.IF(B18="", 0, ((VLOOKUP(D18, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O18+P18))) + Input_AdminFee);
    }

    public double cell_R18() {
        return ExcelFunctions.IF(B18="", 0, (O18+P18-Q18) * Calc_MthlyRate_Guar);
    }

    public double cell_S18() {
        return ExcelFunctions.IF(B18="", 0, Math.max(0, O18+P18-Q18+R18));
    }

    public double cell_A19() {
        return EDATE(A18, 1);
    }

    public double cell_B19() {
        return ExcelFunctions.IF(B18>=Input_ProjYears, "", ExcelFunctions.IF(C18=12, B18+1, B18));
    }

    public double cell_C19() {
        return ExcelFunctions.IF(C18=12, 1, C18+1);
    }

    public double cell_D19() {
        return Input_IssueAge+(B19-1);
    }

    public double cell_F19() {
        return ExcelFunctions.IF(B19="", 0, J18);
    }

    public double cell_G19() {
        return VLOOKUP(A19, Premium!18:1065, 12, TRUE);
    }

    public double cell_H19() {
        return ExcelFunctions.IF(B19="", 0, (VLOOKUP(D19, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F19+G19))) + Input_AdminFee);
    }

    public double cell_I19() {
        return ExcelFunctions.IF(B19="", 0, (F19+G19-H19) * Calc_MthlyRate_Cur);
    }

    public double cell_J19() {
        return ExcelFunctions.IF(B19="", 0, Math.max(0, F19+G19-H19+I19));
    }

    public double cell_K19() {
        return ExcelFunctions.IF(B19="", 0, VLOOKUP(D19, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L19() {
        return ExcelFunctions.IF(B19="", 0, Math.max(0, J19-K19));
    }

    public double cell_M19() {
        return ExcelFunctions.IF(B19="", 0, Math.max(Input_FaceAmt, J19 * VLOOKUP(D19, Table_Rates, 4, FALSE)));
    }

    public double cell_O19() {
        return ACV;
    }

    public double cell_P19() {
        return Premium!L19;
    }

    public double cell_Q19() {
        return ExcelFunctions.IF(B19="", 0, ((VLOOKUP(D19, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O19+P19))) + Input_AdminFee);
    }

    public double cell_R19() {
        return ExcelFunctions.IF(B19="", 0, (O19+P19-Q19) * Calc_MthlyRate_Guar);
    }

    public double cell_S19() {
        return ExcelFunctions.IF(B19="", 0, Math.max(0, O19+P19-Q19+R19));
    }

    public double cell_A20() {
        return EDATE(A19, 1);
    }

    public double cell_B20() {
        return ExcelFunctions.IF(B19>=Input_ProjYears, "", ExcelFunctions.IF(C19=12, B19+1, B19));
    }

    public double cell_C20() {
        return ExcelFunctions.IF(C19=12, 1, C19+1);
    }

    public double cell_D20() {
        return Input_IssueAge+(B20-1);
    }

    public double cell_F20() {
        return ExcelFunctions.IF(B20="", 0, J19);
    }

    public double cell_G20() {
        return VLOOKUP(A20, Premium!19:1066, 12, TRUE);
    }

    public double cell_H20() {
        return ExcelFunctions.IF(B20="", 0, (VLOOKUP(D20, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F20+G20))) + Input_AdminFee);
    }

    public double cell_I20() {
        return ExcelFunctions.IF(B20="", 0, (F20+G20-H20) * Calc_MthlyRate_Cur);
    }

    public double cell_J20() {
        return ExcelFunctions.IF(B20="", 0, Math.max(0, F20+G20-H20+I20));
    }

    public double cell_K20() {
        return ExcelFunctions.IF(B20="", 0, VLOOKUP(D20, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L20() {
        return ExcelFunctions.IF(B20="", 0, Math.max(0, J20-K20));
    }

    public double cell_M20() {
        return ExcelFunctions.IF(B20="", 0, Math.max(Input_FaceAmt, J20 * VLOOKUP(D20, Table_Rates, 4, FALSE)));
    }

    public double cell_O20() {
        return ACV;
    }

    public double cell_P20() {
        return Premium!L20;
    }

    public double cell_Q20() {
        return ExcelFunctions.IF(B20="", 0, ((VLOOKUP(D20, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O20+P20))) + Input_AdminFee);
    }

    public double cell_R20() {
        return ExcelFunctions.IF(B20="", 0, (O20+P20-Q20) * Calc_MthlyRate_Guar);
    }

    public double cell_S20() {
        return ExcelFunctions.IF(B20="", 0, Math.max(0, O20+P20-Q20+R20));
    }

    public double cell_A21() {
        return EDATE(A20, 1);
    }

    public double cell_B21() {
        return ExcelFunctions.IF(B20>=Input_ProjYears, "", ExcelFunctions.IF(C20=12, B20+1, B20));
    }

    public double cell_C21() {
        return ExcelFunctions.IF(C20=12, 1, C20+1);
    }

    public double cell_D21() {
        return Input_IssueAge+(B21-1);
    }

    public double cell_F21() {
        return ExcelFunctions.IF(B21="", 0, J20);
    }

    public double cell_G21() {
        return VLOOKUP(A21, Premium!20:1067, 12, TRUE);
    }

    public double cell_H21() {
        return ExcelFunctions.IF(B21="", 0, (VLOOKUP(D21, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F21+G21))) + Input_AdminFee);
    }

    public double cell_I21() {
        return ExcelFunctions.IF(B21="", 0, (F21+G21-H21) * Calc_MthlyRate_Cur);
    }

    public double cell_J21() {
        return ExcelFunctions.IF(B21="", 0, Math.max(0, F21+G21-H21+I21));
    }

    public double cell_K21() {
        return ExcelFunctions.IF(B21="", 0, VLOOKUP(D21, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L21() {
        return ExcelFunctions.IF(B21="", 0, Math.max(0, J21-K21));
    }

    public double cell_M21() {
        return ExcelFunctions.IF(B21="", 0, Math.max(Input_FaceAmt, J21 * VLOOKUP(D21, Table_Rates, 4, FALSE)));
    }

    public double cell_O21() {
        return ACV;
    }

    public double cell_P21() {
        return Premium!L21;
    }

    public double cell_Q21() {
        return ExcelFunctions.IF(B21="", 0, ((VLOOKUP(D21, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O21+P21))) + Input_AdminFee);
    }

    public double cell_R21() {
        return ExcelFunctions.IF(B21="", 0, (O21+P21-Q21) * Calc_MthlyRate_Guar);
    }

    public double cell_S21() {
        return ExcelFunctions.IF(B21="", 0, Math.max(0, O21+P21-Q21+R21));
    }

    public double cell_A22() {
        return EDATE(A21, 1);
    }

    public double cell_B22() {
        return ExcelFunctions.IF(B21>=Input_ProjYears, "", ExcelFunctions.IF(C21=12, B21+1, B21));
    }

    public double cell_C22() {
        return ExcelFunctions.IF(C21=12, 1, C21+1);
    }

    public double cell_D22() {
        return Input_IssueAge+(B22-1);
    }

    public double cell_F22() {
        return ExcelFunctions.IF(B22="", 0, J21);
    }

    public double cell_G22() {
        return VLOOKUP(A22, Premium!21:1068, 12, TRUE);
    }

    public double cell_H22() {
        return ExcelFunctions.IF(B22="", 0, (VLOOKUP(D22, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F22+G22))) + Input_AdminFee);
    }

    public double cell_I22() {
        return ExcelFunctions.IF(B22="", 0, (F22+G22-H22) * Calc_MthlyRate_Cur);
    }

    public double cell_J22() {
        return ExcelFunctions.IF(B22="", 0, Math.max(0, F22+G22-H22+I22));
    }

    public double cell_K22() {
        return ExcelFunctions.IF(B22="", 0, VLOOKUP(D22, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L22() {
        return ExcelFunctions.IF(B22="", 0, Math.max(0, J22-K22));
    }

    public double cell_M22() {
        return ExcelFunctions.IF(B22="", 0, Math.max(Input_FaceAmt, J22 * VLOOKUP(D22, Table_Rates, 4, FALSE)));
    }

    public double cell_O22() {
        return ACV;
    }

    public double cell_P22() {
        return Premium!L22;
    }

    public double cell_Q22() {
        return ExcelFunctions.IF(B22="", 0, ((VLOOKUP(D22, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O22+P22))) + Input_AdminFee);
    }

    public double cell_R22() {
        return ExcelFunctions.IF(B22="", 0, (O22+P22-Q22) * Calc_MthlyRate_Guar);
    }

    public double cell_S22() {
        return ExcelFunctions.IF(B22="", 0, Math.max(0, O22+P22-Q22+R22));
    }

    public double cell_A23() {
        return EDATE(A22, 1);
    }

    public double cell_B23() {
        return ExcelFunctions.IF(B22>=Input_ProjYears, "", ExcelFunctions.IF(C22=12, B22+1, B22));
    }

    public double cell_C23() {
        return ExcelFunctions.IF(C22=12, 1, C22+1);
    }

    public double cell_D23() {
        return Input_IssueAge+(B23-1);
    }

    public double cell_F23() {
        return ExcelFunctions.IF(B23="", 0, J22);
    }

    public double cell_G23() {
        return VLOOKUP(A23, Premium!22:1069, 12, TRUE);
    }

    public double cell_H23() {
        return ExcelFunctions.IF(B23="", 0, (VLOOKUP(D23, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F23+G23))) + Input_AdminFee);
    }

    public double cell_I23() {
        return ExcelFunctions.IF(B23="", 0, (F23+G23-H23) * Calc_MthlyRate_Cur);
    }

    public double cell_J23() {
        return ExcelFunctions.IF(B23="", 0, Math.max(0, F23+G23-H23+I23));
    }

    public double cell_K23() {
        return ExcelFunctions.IF(B23="", 0, VLOOKUP(D23, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L23() {
        return ExcelFunctions.IF(B23="", 0, Math.max(0, J23-K23));
    }

    public double cell_M23() {
        return ExcelFunctions.IF(B23="", 0, Math.max(Input_FaceAmt, J23 * VLOOKUP(D23, Table_Rates, 4, FALSE)));
    }

    public double cell_O23() {
        return ACV;
    }

    public double cell_P23() {
        return Premium!L23;
    }

    public double cell_Q23() {
        return ExcelFunctions.IF(B23="", 0, ((VLOOKUP(D23, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O23+P23))) + Input_AdminFee);
    }

    public double cell_R23() {
        return ExcelFunctions.IF(B23="", 0, (O23+P23-Q23) * Calc_MthlyRate_Guar);
    }

    public double cell_S23() {
        return ExcelFunctions.IF(B23="", 0, Math.max(0, O23+P23-Q23+R23));
    }

    public double cell_A24() {
        return EDATE(A23, 1);
    }

    public double cell_B24() {
        return ExcelFunctions.IF(B23>=Input_ProjYears, "", ExcelFunctions.IF(C23=12, B23+1, B23));
    }

    public double cell_C24() {
        return ExcelFunctions.IF(C23=12, 1, C23+1);
    }

    public double cell_D24() {
        return Input_IssueAge+(B24-1);
    }

    public double cell_F24() {
        return ExcelFunctions.IF(B24="", 0, J23);
    }

    public double cell_G24() {
        return VLOOKUP(A24, Premium!23:1070, 12, TRUE);
    }

    public double cell_H24() {
        return ExcelFunctions.IF(B24="", 0, (VLOOKUP(D24, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F24+G24))) + Input_AdminFee);
    }

    public double cell_I24() {
        return ExcelFunctions.IF(B24="", 0, (F24+G24-H24) * Calc_MthlyRate_Cur);
    }

    public double cell_J24() {
        return ExcelFunctions.IF(B24="", 0, Math.max(0, F24+G24-H24+I24));
    }

    public double cell_K24() {
        return ExcelFunctions.IF(B24="", 0, VLOOKUP(D24, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L24() {
        return ExcelFunctions.IF(B24="", 0, Math.max(0, J24-K24));
    }

    public double cell_M24() {
        return ExcelFunctions.IF(B24="", 0, Math.max(Input_FaceAmt, J24 * VLOOKUP(D24, Table_Rates, 4, FALSE)));
    }

    public double cell_O24() {
        return ACV;
    }

    public double cell_P24() {
        return Premium!L24;
    }

    public double cell_Q24() {
        return ExcelFunctions.IF(B24="", 0, ((VLOOKUP(D24, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O24+P24))) + Input_AdminFee);
    }

    public double cell_R24() {
        return ExcelFunctions.IF(B24="", 0, (O24+P24-Q24) * Calc_MthlyRate_Guar);
    }

    public double cell_S24() {
        return ExcelFunctions.IF(B24="", 0, Math.max(0, O24+P24-Q24+R24));
    }

    public double cell_A25() {
        return EDATE(A24, 1);
    }

    public double cell_B25() {
        return ExcelFunctions.IF(B24>=Input_ProjYears, "", ExcelFunctions.IF(C24=12, B24+1, B24));
    }

    public double cell_C25() {
        return ExcelFunctions.IF(C24=12, 1, C24+1);
    }

    public double cell_D25() {
        return Input_IssueAge+(B25-1);
    }

    public double cell_F25() {
        return ExcelFunctions.IF(B25="", 0, J24);
    }

    public double cell_G25() {
        return VLOOKUP(A25, Premium!24:1071, 12, TRUE);
    }

    public double cell_H25() {
        return ExcelFunctions.IF(B25="", 0, (VLOOKUP(D25, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F25+G25))) + Input_AdminFee);
    }

    public double cell_I25() {
        return ExcelFunctions.IF(B25="", 0, (F25+G25-H25) * Calc_MthlyRate_Cur);
    }

    public double cell_J25() {
        return ExcelFunctions.IF(B25="", 0, Math.max(0, F25+G25-H25+I25));
    }

    public double cell_K25() {
        return ExcelFunctions.IF(B25="", 0, VLOOKUP(D25, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L25() {
        return ExcelFunctions.IF(B25="", 0, Math.max(0, J25-K25));
    }

    public double cell_M25() {
        return ExcelFunctions.IF(B25="", 0, Math.max(Input_FaceAmt, J25 * VLOOKUP(D25, Table_Rates, 4, FALSE)));
    }

    public double cell_O25() {
        return ACV;
    }

    public double cell_P25() {
        return Premium!L25;
    }

    public double cell_Q25() {
        return ExcelFunctions.IF(B25="", 0, ((VLOOKUP(D25, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O25+P25))) + Input_AdminFee);
    }

    public double cell_R25() {
        return ExcelFunctions.IF(B25="", 0, (O25+P25-Q25) * Calc_MthlyRate_Guar);
    }

    public double cell_S25() {
        return ExcelFunctions.IF(B25="", 0, Math.max(0, O25+P25-Q25+R25));
    }

    public double cell_A26() {
        return EDATE(A25, 1);
    }

    public double cell_B26() {
        return ExcelFunctions.IF(B25>=Input_ProjYears, "", ExcelFunctions.IF(C25=12, B25+1, B25));
    }

    public double cell_C26() {
        return ExcelFunctions.IF(C25=12, 1, C25+1);
    }

    public double cell_D26() {
        return Input_IssueAge+(B26-1);
    }

    public double cell_F26() {
        return ExcelFunctions.IF(B26="", 0, J25);
    }

    public double cell_G26() {
        return VLOOKUP(A26, Premium!25:1072, 12, TRUE);
    }

    public double cell_H26() {
        return ExcelFunctions.IF(B26="", 0, (VLOOKUP(D26, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F26+G26))) + Input_AdminFee);
    }

    public double cell_I26() {
        return ExcelFunctions.IF(B26="", 0, (F26+G26-H26) * Calc_MthlyRate_Cur);
    }

    public double cell_J26() {
        return ExcelFunctions.IF(B26="", 0, Math.max(0, F26+G26-H26+I26));
    }

    public double cell_K26() {
        return ExcelFunctions.IF(B26="", 0, VLOOKUP(D26, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L26() {
        return ExcelFunctions.IF(B26="", 0, Math.max(0, J26-K26));
    }

    public double cell_M26() {
        return ExcelFunctions.IF(B26="", 0, Math.max(Input_FaceAmt, J26 * VLOOKUP(D26, Table_Rates, 4, FALSE)));
    }

    public double cell_O26() {
        return ACV;
    }

    public double cell_P26() {
        return Premium!L26;
    }

    public double cell_Q26() {
        return ExcelFunctions.IF(B26="", 0, ((VLOOKUP(D26, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O26+P26))) + Input_AdminFee);
    }

    public double cell_R26() {
        return ExcelFunctions.IF(B26="", 0, (O26+P26-Q26) * Calc_MthlyRate_Guar);
    }

    public double cell_S26() {
        return ExcelFunctions.IF(B26="", 0, Math.max(0, O26+P26-Q26+R26));
    }

    public double cell_A27() {
        return EDATE(A26, 1);
    }

    public double cell_B27() {
        return ExcelFunctions.IF(B26>=Input_ProjYears, "", ExcelFunctions.IF(C26=12, B26+1, B26));
    }

    public double cell_C27() {
        return ExcelFunctions.IF(C26=12, 1, C26+1);
    }

    public double cell_D27() {
        return Input_IssueAge+(B27-1);
    }

    public double cell_F27() {
        return ExcelFunctions.IF(B27="", 0, J26);
    }

    public double cell_G27() {
        return VLOOKUP(A27, Premium!26:1073, 12, TRUE);
    }

    public double cell_H27() {
        return ExcelFunctions.IF(B27="", 0, (VLOOKUP(D27, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F27+G27))) + Input_AdminFee);
    }

    public double cell_I27() {
        return ExcelFunctions.IF(B27="", 0, (F27+G27-H27) * Calc_MthlyRate_Cur);
    }

    public double cell_J27() {
        return ExcelFunctions.IF(B27="", 0, Math.max(0, F27+G27-H27+I27));
    }

    public double cell_K27() {
        return ExcelFunctions.IF(B27="", 0, VLOOKUP(D27, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L27() {
        return ExcelFunctions.IF(B27="", 0, Math.max(0, J27-K27));
    }

    public double cell_M27() {
        return ExcelFunctions.IF(B27="", 0, Math.max(Input_FaceAmt, J27 * VLOOKUP(D27, Table_Rates, 4, FALSE)));
    }

    public double cell_O27() {
        return ACV;
    }

    public double cell_P27() {
        return Premium!L27;
    }

    public double cell_Q27() {
        return ExcelFunctions.IF(B27="", 0, ((VLOOKUP(D27, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O27+P27))) + Input_AdminFee);
    }

    public double cell_R27() {
        return ExcelFunctions.IF(B27="", 0, (O27+P27-Q27) * Calc_MthlyRate_Guar);
    }

    public double cell_S27() {
        return ExcelFunctions.IF(B27="", 0, Math.max(0, O27+P27-Q27+R27));
    }

    public double cell_A28() {
        return EDATE(A27, 1);
    }

    public double cell_B28() {
        return ExcelFunctions.IF(B27>=Input_ProjYears, "", ExcelFunctions.IF(C27=12, B27+1, B27));
    }

    public double cell_C28() {
        return ExcelFunctions.IF(C27=12, 1, C27+1);
    }

    public double cell_D28() {
        return Input_IssueAge+(B28-1);
    }

    public double cell_F28() {
        return ExcelFunctions.IF(B28="", 0, J27);
    }

    public double cell_G28() {
        return VLOOKUP(A28, Premium!27:1074, 12, TRUE);
    }

    public double cell_H28() {
        return ExcelFunctions.IF(B28="", 0, (VLOOKUP(D28, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F28+G28))) + Input_AdminFee);
    }

    public double cell_I28() {
        return ExcelFunctions.IF(B28="", 0, (F28+G28-H28) * Calc_MthlyRate_Cur);
    }

    public double cell_J28() {
        return ExcelFunctions.IF(B28="", 0, Math.max(0, F28+G28-H28+I28));
    }

    public double cell_K28() {
        return ExcelFunctions.IF(B28="", 0, VLOOKUP(D28, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L28() {
        return ExcelFunctions.IF(B28="", 0, Math.max(0, J28-K28));
    }

    public double cell_M28() {
        return ExcelFunctions.IF(B28="", 0, Math.max(Input_FaceAmt, J28 * VLOOKUP(D28, Table_Rates, 4, FALSE)));
    }

    public double cell_O28() {
        return ACV;
    }

    public double cell_P28() {
        return Premium!L28;
    }

    public double cell_Q28() {
        return ExcelFunctions.IF(B28="", 0, ((VLOOKUP(D28, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O28+P28))) + Input_AdminFee);
    }

    public double cell_R28() {
        return ExcelFunctions.IF(B28="", 0, (O28+P28-Q28) * Calc_MthlyRate_Guar);
    }

    public double cell_S28() {
        return ExcelFunctions.IF(B28="", 0, Math.max(0, O28+P28-Q28+R28));
    }

    public double cell_A29() {
        return EDATE(A28, 1);
    }

    public double cell_B29() {
        return ExcelFunctions.IF(B28>=Input_ProjYears, "", ExcelFunctions.IF(C28=12, B28+1, B28));
    }

    public double cell_C29() {
        return ExcelFunctions.IF(C28=12, 1, C28+1);
    }

    public double cell_D29() {
        return Input_IssueAge+(B29-1);
    }

    public double cell_F29() {
        return ExcelFunctions.IF(B29="", 0, J28);
    }

    public double cell_G29() {
        return VLOOKUP(A29, Premium!28:1075, 12, TRUE);
    }

    public double cell_H29() {
        return ExcelFunctions.IF(B29="", 0, (VLOOKUP(D29, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F29+G29))) + Input_AdminFee);
    }

    public double cell_I29() {
        return ExcelFunctions.IF(B29="", 0, (F29+G29-H29) * Calc_MthlyRate_Cur);
    }

    public double cell_J29() {
        return ExcelFunctions.IF(B29="", 0, Math.max(0, F29+G29-H29+I29));
    }

    public double cell_K29() {
        return ExcelFunctions.IF(B29="", 0, VLOOKUP(D29, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L29() {
        return ExcelFunctions.IF(B29="", 0, Math.max(0, J29-K29));
    }

    public double cell_M29() {
        return ExcelFunctions.IF(B29="", 0, Math.max(Input_FaceAmt, J29 * VLOOKUP(D29, Table_Rates, 4, FALSE)));
    }

    public double cell_O29() {
        return ACV;
    }

    public double cell_P29() {
        return Premium!L29;
    }

    public double cell_Q29() {
        return ExcelFunctions.IF(B29="", 0, ((VLOOKUP(D29, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O29+P29))) + Input_AdminFee);
    }

    public double cell_R29() {
        return ExcelFunctions.IF(B29="", 0, (O29+P29-Q29) * Calc_MthlyRate_Guar);
    }

    public double cell_S29() {
        return ExcelFunctions.IF(B29="", 0, Math.max(0, O29+P29-Q29+R29));
    }

    public double cell_A30() {
        return EDATE(A29, 1);
    }

    public double cell_B30() {
        return ExcelFunctions.IF(B29>=Input_ProjYears, "", ExcelFunctions.IF(C29=12, B29+1, B29));
    }

    public double cell_C30() {
        return ExcelFunctions.IF(C29=12, 1, C29+1);
    }

    public double cell_D30() {
        return Input_IssueAge+(B30-1);
    }

    public double cell_F30() {
        return ExcelFunctions.IF(B30="", 0, J29);
    }

    public double cell_G30() {
        return VLOOKUP(A30, Premium!29:1076, 12, TRUE);
    }

    public double cell_H30() {
        return ExcelFunctions.IF(B30="", 0, (VLOOKUP(D30, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F30+G30))) + Input_AdminFee);
    }

    public double cell_I30() {
        return ExcelFunctions.IF(B30="", 0, (F30+G30-H30) * Calc_MthlyRate_Cur);
    }

    public double cell_J30() {
        return ExcelFunctions.IF(B30="", 0, Math.max(0, F30+G30-H30+I30));
    }

    public double cell_K30() {
        return ExcelFunctions.IF(B30="", 0, VLOOKUP(D30, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L30() {
        return ExcelFunctions.IF(B30="", 0, Math.max(0, J30-K30));
    }

    public double cell_M30() {
        return ExcelFunctions.IF(B30="", 0, Math.max(Input_FaceAmt, J30 * VLOOKUP(D30, Table_Rates, 4, FALSE)));
    }

    public double cell_O30() {
        return ACV;
    }

    public double cell_P30() {
        return Premium!L30;
    }

    public double cell_Q30() {
        return ExcelFunctions.IF(B30="", 0, ((VLOOKUP(D30, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O30+P30))) + Input_AdminFee);
    }

    public double cell_R30() {
        return ExcelFunctions.IF(B30="", 0, (O30+P30-Q30) * Calc_MthlyRate_Guar);
    }

    public double cell_S30() {
        return ExcelFunctions.IF(B30="", 0, Math.max(0, O30+P30-Q30+R30));
    }

    public double cell_A31() {
        return EDATE(A30, 1);
    }

    public double cell_B31() {
        return ExcelFunctions.IF(B30>=Input_ProjYears, "", ExcelFunctions.IF(C30=12, B30+1, B30));
    }

    public double cell_C31() {
        return ExcelFunctions.IF(C30=12, 1, C30+1);
    }

    public double cell_D31() {
        return Input_IssueAge+(B31-1);
    }

    public double cell_F31() {
        return ExcelFunctions.IF(B31="", 0, J30);
    }

    public double cell_G31() {
        return VLOOKUP(A31, Premium!30:1077, 12, TRUE);
    }

    public double cell_H31() {
        return ExcelFunctions.IF(B31="", 0, (VLOOKUP(D31, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F31+G31))) + Input_AdminFee);
    }

    public double cell_I31() {
        return ExcelFunctions.IF(B31="", 0, (F31+G31-H31) * Calc_MthlyRate_Cur);
    }

    public double cell_J31() {
        return ExcelFunctions.IF(B31="", 0, Math.max(0, F31+G31-H31+I31));
    }

    public double cell_K31() {
        return ExcelFunctions.IF(B31="", 0, VLOOKUP(D31, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L31() {
        return ExcelFunctions.IF(B31="", 0, Math.max(0, J31-K31));
    }

    public double cell_M31() {
        return ExcelFunctions.IF(B31="", 0, Math.max(Input_FaceAmt, J31 * VLOOKUP(D31, Table_Rates, 4, FALSE)));
    }

    public double cell_O31() {
        return ACV;
    }

    public double cell_P31() {
        return Premium!L31;
    }

    public double cell_Q31() {
        return ExcelFunctions.IF(B31="", 0, ((VLOOKUP(D31, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O31+P31))) + Input_AdminFee);
    }

    public double cell_R31() {
        return ExcelFunctions.IF(B31="", 0, (O31+P31-Q31) * Calc_MthlyRate_Guar);
    }

    public double cell_S31() {
        return ExcelFunctions.IF(B31="", 0, Math.max(0, O31+P31-Q31+R31));
    }

    public double cell_A32() {
        return EDATE(A31, 1);
    }

    public double cell_B32() {
        return ExcelFunctions.IF(B31>=Input_ProjYears, "", ExcelFunctions.IF(C31=12, B31+1, B31));
    }

    public double cell_C32() {
        return ExcelFunctions.IF(C31=12, 1, C31+1);
    }

    public double cell_D32() {
        return Input_IssueAge+(B32-1);
    }

    public double cell_F32() {
        return ExcelFunctions.IF(B32="", 0, J31);
    }

    public double cell_G32() {
        return VLOOKUP(A32, Premium!31:1078, 12, TRUE);
    }

    public double cell_H32() {
        return ExcelFunctions.IF(B32="", 0, (VLOOKUP(D32, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F32+G32))) + Input_AdminFee);
    }

    public double cell_I32() {
        return ExcelFunctions.IF(B32="", 0, (F32+G32-H32) * Calc_MthlyRate_Cur);
    }

    public double cell_J32() {
        return ExcelFunctions.IF(B32="", 0, Math.max(0, F32+G32-H32+I32));
    }

    public double cell_K32() {
        return ExcelFunctions.IF(B32="", 0, VLOOKUP(D32, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L32() {
        return ExcelFunctions.IF(B32="", 0, Math.max(0, J32-K32));
    }

    public double cell_M32() {
        return ExcelFunctions.IF(B32="", 0, Math.max(Input_FaceAmt, J32 * VLOOKUP(D32, Table_Rates, 4, FALSE)));
    }

    public double cell_O32() {
        return ACV;
    }

    public double cell_P32() {
        return Premium!L32;
    }

    public double cell_Q32() {
        return ExcelFunctions.IF(B32="", 0, ((VLOOKUP(D32, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O32+P32))) + Input_AdminFee);
    }

    public double cell_R32() {
        return ExcelFunctions.IF(B32="", 0, (O32+P32-Q32) * Calc_MthlyRate_Guar);
    }

    public double cell_S32() {
        return ExcelFunctions.IF(B32="", 0, Math.max(0, O32+P32-Q32+R32));
    }

    public double cell_A33() {
        return EDATE(A32, 1);
    }

    public double cell_B33() {
        return ExcelFunctions.IF(B32>=Input_ProjYears, "", ExcelFunctions.IF(C32=12, B32+1, B32));
    }

    public double cell_C33() {
        return ExcelFunctions.IF(C32=12, 1, C32+1);
    }

    public double cell_D33() {
        return Input_IssueAge+(B33-1);
    }

    public double cell_F33() {
        return ExcelFunctions.IF(B33="", 0, J32);
    }

    public double cell_G33() {
        return VLOOKUP(A33, Premium!32:1079, 12, TRUE);
    }

    public double cell_H33() {
        return ExcelFunctions.IF(B33="", 0, (VLOOKUP(D33, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F33+G33))) + Input_AdminFee);
    }

    public double cell_I33() {
        return ExcelFunctions.IF(B33="", 0, (F33+G33-H33) * Calc_MthlyRate_Cur);
    }

    public double cell_J33() {
        return ExcelFunctions.IF(B33="", 0, Math.max(0, F33+G33-H33+I33));
    }

    public double cell_K33() {
        return ExcelFunctions.IF(B33="", 0, VLOOKUP(D33, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L33() {
        return ExcelFunctions.IF(B33="", 0, Math.max(0, J33-K33));
    }

    public double cell_M33() {
        return ExcelFunctions.IF(B33="", 0, Math.max(Input_FaceAmt, J33 * VLOOKUP(D33, Table_Rates, 4, FALSE)));
    }

    public double cell_O33() {
        return ACV;
    }

    public double cell_P33() {
        return Premium!L33;
    }

    public double cell_Q33() {
        return ExcelFunctions.IF(B33="", 0, ((VLOOKUP(D33, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O33+P33))) + Input_AdminFee);
    }

    public double cell_R33() {
        return ExcelFunctions.IF(B33="", 0, (O33+P33-Q33) * Calc_MthlyRate_Guar);
    }

    public double cell_S33() {
        return ExcelFunctions.IF(B33="", 0, Math.max(0, O33+P33-Q33+R33));
    }

    public double cell_A34() {
        return EDATE(A33, 1);
    }

    public double cell_B34() {
        return ExcelFunctions.IF(B33>=Input_ProjYears, "", ExcelFunctions.IF(C33=12, B33+1, B33));
    }

    public double cell_C34() {
        return ExcelFunctions.IF(C33=12, 1, C33+1);
    }

    public double cell_D34() {
        return Input_IssueAge+(B34-1);
    }

    public double cell_F34() {
        return ExcelFunctions.IF(B34="", 0, J33);
    }

    public double cell_G34() {
        return VLOOKUP(A34, Premium!33:1080, 12, TRUE);
    }

    public double cell_H34() {
        return ExcelFunctions.IF(B34="", 0, (VLOOKUP(D34, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F34+G34))) + Input_AdminFee);
    }

    public double cell_I34() {
        return ExcelFunctions.IF(B34="", 0, (F34+G34-H34) * Calc_MthlyRate_Cur);
    }

    public double cell_J34() {
        return ExcelFunctions.IF(B34="", 0, Math.max(0, F34+G34-H34+I34));
    }

    public double cell_K34() {
        return ExcelFunctions.IF(B34="", 0, VLOOKUP(D34, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L34() {
        return ExcelFunctions.IF(B34="", 0, Math.max(0, J34-K34));
    }

    public double cell_M34() {
        return ExcelFunctions.IF(B34="", 0, Math.max(Input_FaceAmt, J34 * VLOOKUP(D34, Table_Rates, 4, FALSE)));
    }

    public double cell_O34() {
        return ACV;
    }

    public double cell_P34() {
        return Premium!L34;
    }

    public double cell_Q34() {
        return ExcelFunctions.IF(B34="", 0, ((VLOOKUP(D34, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O34+P34))) + Input_AdminFee);
    }

    public double cell_R34() {
        return ExcelFunctions.IF(B34="", 0, (O34+P34-Q34) * Calc_MthlyRate_Guar);
    }

    public double cell_S34() {
        return ExcelFunctions.IF(B34="", 0, Math.max(0, O34+P34-Q34+R34));
    }

    public double cell_A35() {
        return EDATE(A34, 1);
    }

    public double cell_B35() {
        return ExcelFunctions.IF(B34>=Input_ProjYears, "", ExcelFunctions.IF(C34=12, B34+1, B34));
    }

    public double cell_C35() {
        return ExcelFunctions.IF(C34=12, 1, C34+1);
    }

    public double cell_D35() {
        return Input_IssueAge+(B35-1);
    }

    public double cell_F35() {
        return ExcelFunctions.IF(B35="", 0, J34);
    }

    public double cell_G35() {
        return VLOOKUP(A35, Premium!34:1081, 12, TRUE);
    }

    public double cell_H35() {
        return ExcelFunctions.IF(B35="", 0, (VLOOKUP(D35, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F35+G35))) + Input_AdminFee);
    }

    public double cell_I35() {
        return ExcelFunctions.IF(B35="", 0, (F35+G35-H35) * Calc_MthlyRate_Cur);
    }

    public double cell_J35() {
        return ExcelFunctions.IF(B35="", 0, Math.max(0, F35+G35-H35+I35));
    }

    public double cell_K35() {
        return ExcelFunctions.IF(B35="", 0, VLOOKUP(D35, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L35() {
        return ExcelFunctions.IF(B35="", 0, Math.max(0, J35-K35));
    }

    public double cell_M35() {
        return ExcelFunctions.IF(B35="", 0, Math.max(Input_FaceAmt, J35 * VLOOKUP(D35, Table_Rates, 4, FALSE)));
    }

    public double cell_O35() {
        return ACV;
    }

    public double cell_P35() {
        return Premium!L35;
    }

    public double cell_Q35() {
        return ExcelFunctions.IF(B35="", 0, ((VLOOKUP(D35, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O35+P35))) + Input_AdminFee);
    }

    public double cell_R35() {
        return ExcelFunctions.IF(B35="", 0, (O35+P35-Q35) * Calc_MthlyRate_Guar);
    }

    public double cell_S35() {
        return ExcelFunctions.IF(B35="", 0, Math.max(0, O35+P35-Q35+R35));
    }

    public double cell_A36() {
        return EDATE(A35, 1);
    }

    public double cell_B36() {
        return ExcelFunctions.IF(B35>=Input_ProjYears, "", ExcelFunctions.IF(C35=12, B35+1, B35));
    }

    public double cell_C36() {
        return ExcelFunctions.IF(C35=12, 1, C35+1);
    }

    public double cell_D36() {
        return Input_IssueAge+(B36-1);
    }

    public double cell_F36() {
        return ExcelFunctions.IF(B36="", 0, J35);
    }

    public double cell_G36() {
        return VLOOKUP(A36, Premium!35:1082, 12, TRUE);
    }

    public double cell_H36() {
        return ExcelFunctions.IF(B36="", 0, (VLOOKUP(D36, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F36+G36))) + Input_AdminFee);
    }

    public double cell_I36() {
        return ExcelFunctions.IF(B36="", 0, (F36+G36-H36) * Calc_MthlyRate_Cur);
    }

    public double cell_J36() {
        return ExcelFunctions.IF(B36="", 0, Math.max(0, F36+G36-H36+I36));
    }

    public double cell_K36() {
        return ExcelFunctions.IF(B36="", 0, VLOOKUP(D36, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L36() {
        return ExcelFunctions.IF(B36="", 0, Math.max(0, J36-K36));
    }

    public double cell_M36() {
        return ExcelFunctions.IF(B36="", 0, Math.max(Input_FaceAmt, J36 * VLOOKUP(D36, Table_Rates, 4, FALSE)));
    }

    public double cell_O36() {
        return ACV;
    }

    public double cell_P36() {
        return Premium!L36;
    }

    public double cell_Q36() {
        return ExcelFunctions.IF(B36="", 0, ((VLOOKUP(D36, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O36+P36))) + Input_AdminFee);
    }

    public double cell_R36() {
        return ExcelFunctions.IF(B36="", 0, (O36+P36-Q36) * Calc_MthlyRate_Guar);
    }

    public double cell_S36() {
        return ExcelFunctions.IF(B36="", 0, Math.max(0, O36+P36-Q36+R36));
    }

    public double cell_A37() {
        return EDATE(A36, 1);
    }

    public double cell_B37() {
        return ExcelFunctions.IF(B36>=Input_ProjYears, "", ExcelFunctions.IF(C36=12, B36+1, B36));
    }

    public double cell_C37() {
        return ExcelFunctions.IF(C36=12, 1, C36+1);
    }

    public double cell_D37() {
        return Input_IssueAge+(B37-1);
    }

    public double cell_F37() {
        return ExcelFunctions.IF(B37="", 0, J36);
    }

    public double cell_G37() {
        return VLOOKUP(A37, Premium!36:1083, 12, TRUE);
    }

    public double cell_H37() {
        return ExcelFunctions.IF(B37="", 0, (VLOOKUP(D37, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F37+G37))) + Input_AdminFee);
    }

    public double cell_I37() {
        return ExcelFunctions.IF(B37="", 0, (F37+G37-H37) * Calc_MthlyRate_Cur);
    }

    public double cell_J37() {
        return ExcelFunctions.IF(B37="", 0, Math.max(0, F37+G37-H37+I37));
    }

    public double cell_K37() {
        return ExcelFunctions.IF(B37="", 0, VLOOKUP(D37, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L37() {
        return ExcelFunctions.IF(B37="", 0, Math.max(0, J37-K37));
    }

    public double cell_M37() {
        return ExcelFunctions.IF(B37="", 0, Math.max(Input_FaceAmt, J37 * VLOOKUP(D37, Table_Rates, 4, FALSE)));
    }

    public double cell_O37() {
        return ACV;
    }

    public double cell_P37() {
        return Premium!L37;
    }

    public double cell_Q37() {
        return ExcelFunctions.IF(B37="", 0, ((VLOOKUP(D37, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O37+P37))) + Input_AdminFee);
    }

    public double cell_R37() {
        return ExcelFunctions.IF(B37="", 0, (O37+P37-Q37) * Calc_MthlyRate_Guar);
    }

    public double cell_S37() {
        return ExcelFunctions.IF(B37="", 0, Math.max(0, O37+P37-Q37+R37));
    }

    public double cell_A38() {
        return EDATE(A37, 1);
    }

    public double cell_B38() {
        return ExcelFunctions.IF(B37>=Input_ProjYears, "", ExcelFunctions.IF(C37=12, B37+1, B37));
    }

    public double cell_C38() {
        return ExcelFunctions.IF(C37=12, 1, C37+1);
    }

    public double cell_D38() {
        return Input_IssueAge+(B38-1);
    }

    public double cell_F38() {
        return ExcelFunctions.IF(B38="", 0, J37);
    }

    public double cell_G38() {
        return VLOOKUP(A38, Premium!37:1084, 12, TRUE);
    }

    public double cell_H38() {
        return ExcelFunctions.IF(B38="", 0, (VLOOKUP(D38, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F38+G38))) + Input_AdminFee);
    }

    public double cell_I38() {
        return ExcelFunctions.IF(B38="", 0, (F38+G38-H38) * Calc_MthlyRate_Cur);
    }

    public double cell_J38() {
        return ExcelFunctions.IF(B38="", 0, Math.max(0, F38+G38-H38+I38));
    }

    public double cell_K38() {
        return ExcelFunctions.IF(B38="", 0, VLOOKUP(D38, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L38() {
        return ExcelFunctions.IF(B38="", 0, Math.max(0, J38-K38));
    }

    public double cell_M38() {
        return ExcelFunctions.IF(B38="", 0, Math.max(Input_FaceAmt, J38 * VLOOKUP(D38, Table_Rates, 4, FALSE)));
    }

    public double cell_O38() {
        return ACV;
    }

    public double cell_P38() {
        return Premium!L38;
    }

    public double cell_Q38() {
        return ExcelFunctions.IF(B38="", 0, ((VLOOKUP(D38, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O38+P38))) + Input_AdminFee);
    }

    public double cell_R38() {
        return ExcelFunctions.IF(B38="", 0, (O38+P38-Q38) * Calc_MthlyRate_Guar);
    }

    public double cell_S38() {
        return ExcelFunctions.IF(B38="", 0, Math.max(0, O38+P38-Q38+R38));
    }

    public double cell_A39() {
        return EDATE(A38, 1);
    }

    public double cell_B39() {
        return ExcelFunctions.IF(B38>=Input_ProjYears, "", ExcelFunctions.IF(C38=12, B38+1, B38));
    }

    public double cell_C39() {
        return ExcelFunctions.IF(C38=12, 1, C38+1);
    }

    public double cell_D39() {
        return Input_IssueAge+(B39-1);
    }

    public double cell_F39() {
        return ExcelFunctions.IF(B39="", 0, J38);
    }

    public double cell_G39() {
        return VLOOKUP(A39, Premium!38:1085, 12, TRUE);
    }

    public double cell_H39() {
        return ExcelFunctions.IF(B39="", 0, (VLOOKUP(D39, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F39+G39))) + Input_AdminFee);
    }

    public double cell_I39() {
        return ExcelFunctions.IF(B39="", 0, (F39+G39-H39) * Calc_MthlyRate_Cur);
    }

    public double cell_J39() {
        return ExcelFunctions.IF(B39="", 0, Math.max(0, F39+G39-H39+I39));
    }

    public double cell_K39() {
        return ExcelFunctions.IF(B39="", 0, VLOOKUP(D39, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L39() {
        return ExcelFunctions.IF(B39="", 0, Math.max(0, J39-K39));
    }

    public double cell_M39() {
        return ExcelFunctions.IF(B39="", 0, Math.max(Input_FaceAmt, J39 * VLOOKUP(D39, Table_Rates, 4, FALSE)));
    }

    public double cell_O39() {
        return ACV;
    }

    public double cell_P39() {
        return Premium!L39;
    }

    public double cell_Q39() {
        return ExcelFunctions.IF(B39="", 0, ((VLOOKUP(D39, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O39+P39))) + Input_AdminFee);
    }

    public double cell_R39() {
        return ExcelFunctions.IF(B39="", 0, (O39+P39-Q39) * Calc_MthlyRate_Guar);
    }

    public double cell_S39() {
        return ExcelFunctions.IF(B39="", 0, Math.max(0, O39+P39-Q39+R39));
    }

    public double cell_A40() {
        return EDATE(A39, 1);
    }

    public double cell_B40() {
        return ExcelFunctions.IF(B39>=Input_ProjYears, "", ExcelFunctions.IF(C39=12, B39+1, B39));
    }

    public double cell_C40() {
        return ExcelFunctions.IF(C39=12, 1, C39+1);
    }

    public double cell_D40() {
        return Input_IssueAge+(B40-1);
    }

    public double cell_F40() {
        return ExcelFunctions.IF(B40="", 0, J39);
    }

    public double cell_G40() {
        return VLOOKUP(A40, Premium!39:1086, 12, TRUE);
    }

    public double cell_H40() {
        return ExcelFunctions.IF(B40="", 0, (VLOOKUP(D40, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F40+G40))) + Input_AdminFee);
    }

    public double cell_I40() {
        return ExcelFunctions.IF(B40="", 0, (F40+G40-H40) * Calc_MthlyRate_Cur);
    }

    public double cell_J40() {
        return ExcelFunctions.IF(B40="", 0, Math.max(0, F40+G40-H40+I40));
    }

    public double cell_K40() {
        return ExcelFunctions.IF(B40="", 0, VLOOKUP(D40, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L40() {
        return ExcelFunctions.IF(B40="", 0, Math.max(0, J40-K40));
    }

    public double cell_M40() {
        return ExcelFunctions.IF(B40="", 0, Math.max(Input_FaceAmt, J40 * VLOOKUP(D40, Table_Rates, 4, FALSE)));
    }

    public double cell_O40() {
        return ACV;
    }

    public double cell_P40() {
        return Premium!L40;
    }

    public double cell_Q40() {
        return ExcelFunctions.IF(B40="", 0, ((VLOOKUP(D40, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O40+P40))) + Input_AdminFee);
    }

    public double cell_R40() {
        return ExcelFunctions.IF(B40="", 0, (O40+P40-Q40) * Calc_MthlyRate_Guar);
    }

    public double cell_S40() {
        return ExcelFunctions.IF(B40="", 0, Math.max(0, O40+P40-Q40+R40));
    }

    public double cell_A41() {
        return EDATE(A40, 1);
    }

    public double cell_B41() {
        return ExcelFunctions.IF(B40>=Input_ProjYears, "", ExcelFunctions.IF(C40=12, B40+1, B40));
    }

    public double cell_C41() {
        return ExcelFunctions.IF(C40=12, 1, C40+1);
    }

    public double cell_D41() {
        return Input_IssueAge+(B41-1);
    }

    public double cell_F41() {
        return ExcelFunctions.IF(B41="", 0, J40);
    }

    public double cell_G41() {
        return VLOOKUP(A41, Premium!40:1087, 12, TRUE);
    }

    public double cell_H41() {
        return ExcelFunctions.IF(B41="", 0, (VLOOKUP(D41, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F41+G41))) + Input_AdminFee);
    }

    public double cell_I41() {
        return ExcelFunctions.IF(B41="", 0, (F41+G41-H41) * Calc_MthlyRate_Cur);
    }

    public double cell_J41() {
        return ExcelFunctions.IF(B41="", 0, Math.max(0, F41+G41-H41+I41));
    }

    public double cell_K41() {
        return ExcelFunctions.IF(B41="", 0, VLOOKUP(D41, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L41() {
        return ExcelFunctions.IF(B41="", 0, Math.max(0, J41-K41));
    }

    public double cell_M41() {
        return ExcelFunctions.IF(B41="", 0, Math.max(Input_FaceAmt, J41 * VLOOKUP(D41, Table_Rates, 4, FALSE)));
    }

    public double cell_O41() {
        return ACV;
    }

    public double cell_P41() {
        return Premium!L41;
    }

    public double cell_Q41() {
        return ExcelFunctions.IF(B41="", 0, ((VLOOKUP(D41, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O41+P41))) + Input_AdminFee);
    }

    public double cell_R41() {
        return ExcelFunctions.IF(B41="", 0, (O41+P41-Q41) * Calc_MthlyRate_Guar);
    }

    public double cell_S41() {
        return ExcelFunctions.IF(B41="", 0, Math.max(0, O41+P41-Q41+R41));
    }

    public double cell_A42() {
        return EDATE(A41, 1);
    }

    public double cell_B42() {
        return ExcelFunctions.IF(B41>=Input_ProjYears, "", ExcelFunctions.IF(C41=12, B41+1, B41));
    }

    public double cell_C42() {
        return ExcelFunctions.IF(C41=12, 1, C41+1);
    }

    public double cell_D42() {
        return Input_IssueAge+(B42-1);
    }

    public double cell_F42() {
        return ExcelFunctions.IF(B42="", 0, J41);
    }

    public double cell_G42() {
        return VLOOKUP(A42, Premium!41:1088, 12, TRUE);
    }

    public double cell_H42() {
        return ExcelFunctions.IF(B42="", 0, (VLOOKUP(D42, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F42+G42))) + Input_AdminFee);
    }

    public double cell_I42() {
        return ExcelFunctions.IF(B42="", 0, (F42+G42-H42) * Calc_MthlyRate_Cur);
    }

    public double cell_J42() {
        return ExcelFunctions.IF(B42="", 0, Math.max(0, F42+G42-H42+I42));
    }

    public double cell_K42() {
        return ExcelFunctions.IF(B42="", 0, VLOOKUP(D42, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L42() {
        return ExcelFunctions.IF(B42="", 0, Math.max(0, J42-K42));
    }

    public double cell_M42() {
        return ExcelFunctions.IF(B42="", 0, Math.max(Input_FaceAmt, J42 * VLOOKUP(D42, Table_Rates, 4, FALSE)));
    }

    public double cell_O42() {
        return ACV;
    }

    public double cell_P42() {
        return Premium!L42;
    }

    public double cell_Q42() {
        return ExcelFunctions.IF(B42="", 0, ((VLOOKUP(D42, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O42+P42))) + Input_AdminFee);
    }

    public double cell_R42() {
        return ExcelFunctions.IF(B42="", 0, (O42+P42-Q42) * Calc_MthlyRate_Guar);
    }

    public double cell_S42() {
        return ExcelFunctions.IF(B42="", 0, Math.max(0, O42+P42-Q42+R42));
    }

    public double cell_A43() {
        return EDATE(A42, 1);
    }

    public double cell_B43() {
        return ExcelFunctions.IF(B42>=Input_ProjYears, "", ExcelFunctions.IF(C42=12, B42+1, B42));
    }

    public double cell_C43() {
        return ExcelFunctions.IF(C42=12, 1, C42+1);
    }

    public double cell_D43() {
        return Input_IssueAge+(B43-1);
    }

    public double cell_F43() {
        return ExcelFunctions.IF(B43="", 0, J42);
    }

    public double cell_G43() {
        return VLOOKUP(A43, Premium!42:1089, 12, TRUE);
    }

    public double cell_H43() {
        return ExcelFunctions.IF(B43="", 0, (VLOOKUP(D43, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F43+G43))) + Input_AdminFee);
    }

    public double cell_I43() {
        return ExcelFunctions.IF(B43="", 0, (F43+G43-H43) * Calc_MthlyRate_Cur);
    }

    public double cell_J43() {
        return ExcelFunctions.IF(B43="", 0, Math.max(0, F43+G43-H43+I43));
    }

    public double cell_K43() {
        return ExcelFunctions.IF(B43="", 0, VLOOKUP(D43, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L43() {
        return ExcelFunctions.IF(B43="", 0, Math.max(0, J43-K43));
    }

    public double cell_M43() {
        return ExcelFunctions.IF(B43="", 0, Math.max(Input_FaceAmt, J43 * VLOOKUP(D43, Table_Rates, 4, FALSE)));
    }

    public double cell_O43() {
        return ACV;
    }

    public double cell_P43() {
        return Premium!L43;
    }

    public double cell_Q43() {
        return ExcelFunctions.IF(B43="", 0, ((VLOOKUP(D43, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O43+P43))) + Input_AdminFee);
    }

    public double cell_R43() {
        return ExcelFunctions.IF(B43="", 0, (O43+P43-Q43) * Calc_MthlyRate_Guar);
    }

    public double cell_S43() {
        return ExcelFunctions.IF(B43="", 0, Math.max(0, O43+P43-Q43+R43));
    }

    public double cell_A44() {
        return EDATE(A43, 1);
    }

    public double cell_B44() {
        return ExcelFunctions.IF(B43>=Input_ProjYears, "", ExcelFunctions.IF(C43=12, B43+1, B43));
    }

    public double cell_C44() {
        return ExcelFunctions.IF(C43=12, 1, C43+1);
    }

    public double cell_D44() {
        return Input_IssueAge+(B44-1);
    }

    public double cell_F44() {
        return ExcelFunctions.IF(B44="", 0, J43);
    }

    public double cell_G44() {
        return VLOOKUP(A44, Premium!43:1090, 12, TRUE);
    }

    public double cell_H44() {
        return ExcelFunctions.IF(B44="", 0, (VLOOKUP(D44, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F44+G44))) + Input_AdminFee);
    }

    public double cell_I44() {
        return ExcelFunctions.IF(B44="", 0, (F44+G44-H44) * Calc_MthlyRate_Cur);
    }

    public double cell_J44() {
        return ExcelFunctions.IF(B44="", 0, Math.max(0, F44+G44-H44+I44));
    }

    public double cell_K44() {
        return ExcelFunctions.IF(B44="", 0, VLOOKUP(D44, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L44() {
        return ExcelFunctions.IF(B44="", 0, Math.max(0, J44-K44));
    }

    public double cell_M44() {
        return ExcelFunctions.IF(B44="", 0, Math.max(Input_FaceAmt, J44 * VLOOKUP(D44, Table_Rates, 4, FALSE)));
    }

    public double cell_O44() {
        return ACV;
    }

    public double cell_P44() {
        return Premium!L44;
    }

    public double cell_Q44() {
        return ExcelFunctions.IF(B44="", 0, ((VLOOKUP(D44, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O44+P44))) + Input_AdminFee);
    }

    public double cell_R44() {
        return ExcelFunctions.IF(B44="", 0, (O44+P44-Q44) * Calc_MthlyRate_Guar);
    }

    public double cell_S44() {
        return ExcelFunctions.IF(B44="", 0, Math.max(0, O44+P44-Q44+R44));
    }

    public double cell_A45() {
        return EDATE(A44, 1);
    }

    public double cell_B45() {
        return ExcelFunctions.IF(B44>=Input_ProjYears, "", ExcelFunctions.IF(C44=12, B44+1, B44));
    }

    public double cell_C45() {
        return ExcelFunctions.IF(C44=12, 1, C44+1);
    }

    public double cell_D45() {
        return Input_IssueAge+(B45-1);
    }

    public double cell_F45() {
        return ExcelFunctions.IF(B45="", 0, J44);
    }

    public double cell_G45() {
        return VLOOKUP(A45, Premium!44:1091, 12, TRUE);
    }

    public double cell_H45() {
        return ExcelFunctions.IF(B45="", 0, (VLOOKUP(D45, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F45+G45))) + Input_AdminFee);
    }

    public double cell_I45() {
        return ExcelFunctions.IF(B45="", 0, (F45+G45-H45) * Calc_MthlyRate_Cur);
    }

    public double cell_J45() {
        return ExcelFunctions.IF(B45="", 0, Math.max(0, F45+G45-H45+I45));
    }

    public double cell_K45() {
        return ExcelFunctions.IF(B45="", 0, VLOOKUP(D45, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L45() {
        return ExcelFunctions.IF(B45="", 0, Math.max(0, J45-K45));
    }

    public double cell_M45() {
        return ExcelFunctions.IF(B45="", 0, Math.max(Input_FaceAmt, J45 * VLOOKUP(D45, Table_Rates, 4, FALSE)));
    }

    public double cell_O45() {
        return ACV;
    }

    public double cell_P45() {
        return Premium!L45;
    }

    public double cell_Q45() {
        return ExcelFunctions.IF(B45="", 0, ((VLOOKUP(D45, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O45+P45))) + Input_AdminFee);
    }

    public double cell_R45() {
        return ExcelFunctions.IF(B45="", 0, (O45+P45-Q45) * Calc_MthlyRate_Guar);
    }

    public double cell_S45() {
        return ExcelFunctions.IF(B45="", 0, Math.max(0, O45+P45-Q45+R45));
    }

    public double cell_A46() {
        return EDATE(A45, 1);
    }

    public double cell_B46() {
        return ExcelFunctions.IF(B45>=Input_ProjYears, "", ExcelFunctions.IF(C45=12, B45+1, B45));
    }

    public double cell_C46() {
        return ExcelFunctions.IF(C45=12, 1, C45+1);
    }

    public double cell_D46() {
        return Input_IssueAge+(B46-1);
    }

    public double cell_F46() {
        return ExcelFunctions.IF(B46="", 0, J45);
    }

    public double cell_G46() {
        return VLOOKUP(A46, Premium!45:1092, 12, TRUE);
    }

    public double cell_H46() {
        return ExcelFunctions.IF(B46="", 0, (VLOOKUP(D46, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F46+G46))) + Input_AdminFee);
    }

    public double cell_I46() {
        return ExcelFunctions.IF(B46="", 0, (F46+G46-H46) * Calc_MthlyRate_Cur);
    }

    public double cell_J46() {
        return ExcelFunctions.IF(B46="", 0, Math.max(0, F46+G46-H46+I46));
    }

    public double cell_K46() {
        return ExcelFunctions.IF(B46="", 0, VLOOKUP(D46, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L46() {
        return ExcelFunctions.IF(B46="", 0, Math.max(0, J46-K46));
    }

    public double cell_M46() {
        return ExcelFunctions.IF(B46="", 0, Math.max(Input_FaceAmt, J46 * VLOOKUP(D46, Table_Rates, 4, FALSE)));
    }

    public double cell_O46() {
        return ACV;
    }

    public double cell_P46() {
        return Premium!L46;
    }

    public double cell_Q46() {
        return ExcelFunctions.IF(B46="", 0, ((VLOOKUP(D46, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O46+P46))) + Input_AdminFee);
    }

    public double cell_R46() {
        return ExcelFunctions.IF(B46="", 0, (O46+P46-Q46) * Calc_MthlyRate_Guar);
    }

    public double cell_S46() {
        return ExcelFunctions.IF(B46="", 0, Math.max(0, O46+P46-Q46+R46));
    }

    public double cell_A47() {
        return EDATE(A46, 1);
    }

    public double cell_B47() {
        return ExcelFunctions.IF(B46>=Input_ProjYears, "", ExcelFunctions.IF(C46=12, B46+1, B46));
    }

    public double cell_C47() {
        return ExcelFunctions.IF(C46=12, 1, C46+1);
    }

    public double cell_D47() {
        return Input_IssueAge+(B47-1);
    }

    public double cell_F47() {
        return ExcelFunctions.IF(B47="", 0, J46);
    }

    public double cell_G47() {
        return VLOOKUP(A47, Premium!46:1093, 12, TRUE);
    }

    public double cell_H47() {
        return ExcelFunctions.IF(B47="", 0, (VLOOKUP(D47, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F47+G47))) + Input_AdminFee);
    }

    public double cell_I47() {
        return ExcelFunctions.IF(B47="", 0, (F47+G47-H47) * Calc_MthlyRate_Cur);
    }

    public double cell_J47() {
        return ExcelFunctions.IF(B47="", 0, Math.max(0, F47+G47-H47+I47));
    }

    public double cell_K47() {
        return ExcelFunctions.IF(B47="", 0, VLOOKUP(D47, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L47() {
        return ExcelFunctions.IF(B47="", 0, Math.max(0, J47-K47));
    }

    public double cell_M47() {
        return ExcelFunctions.IF(B47="", 0, Math.max(Input_FaceAmt, J47 * VLOOKUP(D47, Table_Rates, 4, FALSE)));
    }

    public double cell_O47() {
        return ACV;
    }

    public double cell_P47() {
        return Premium!L47;
    }

    public double cell_Q47() {
        return ExcelFunctions.IF(B47="", 0, ((VLOOKUP(D47, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O47+P47))) + Input_AdminFee);
    }

    public double cell_R47() {
        return ExcelFunctions.IF(B47="", 0, (O47+P47-Q47) * Calc_MthlyRate_Guar);
    }

    public double cell_S47() {
        return ExcelFunctions.IF(B47="", 0, Math.max(0, O47+P47-Q47+R47));
    }

    public double cell_A48() {
        return EDATE(A47, 1);
    }

    public double cell_B48() {
        return ExcelFunctions.IF(B47>=Input_ProjYears, "", ExcelFunctions.IF(C47=12, B47+1, B47));
    }

    public double cell_C48() {
        return ExcelFunctions.IF(C47=12, 1, C47+1);
    }

    public double cell_D48() {
        return Input_IssueAge+(B48-1);
    }

    public double cell_F48() {
        return ExcelFunctions.IF(B48="", 0, J47);
    }

    public double cell_G48() {
        return VLOOKUP(A48, Premium!47:1094, 12, TRUE);
    }

    public double cell_H48() {
        return ExcelFunctions.IF(B48="", 0, (VLOOKUP(D48, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F48+G48))) + Input_AdminFee);
    }

    public double cell_I48() {
        return ExcelFunctions.IF(B48="", 0, (F48+G48-H48) * Calc_MthlyRate_Cur);
    }

    public double cell_J48() {
        return ExcelFunctions.IF(B48="", 0, Math.max(0, F48+G48-H48+I48));
    }

    public double cell_K48() {
        return ExcelFunctions.IF(B48="", 0, VLOOKUP(D48, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L48() {
        return ExcelFunctions.IF(B48="", 0, Math.max(0, J48-K48));
    }

    public double cell_M48() {
        return ExcelFunctions.IF(B48="", 0, Math.max(Input_FaceAmt, J48 * VLOOKUP(D48, Table_Rates, 4, FALSE)));
    }

    public double cell_O48() {
        return ACV;
    }

    public double cell_P48() {
        return Premium!L48;
    }

    public double cell_Q48() {
        return ExcelFunctions.IF(B48="", 0, ((VLOOKUP(D48, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O48+P48))) + Input_AdminFee);
    }

    public double cell_R48() {
        return ExcelFunctions.IF(B48="", 0, (O48+P48-Q48) * Calc_MthlyRate_Guar);
    }

    public double cell_S48() {
        return ExcelFunctions.IF(B48="", 0, Math.max(0, O48+P48-Q48+R48));
    }

    public double cell_A49() {
        return EDATE(A48, 1);
    }

    public double cell_B49() {
        return ExcelFunctions.IF(B48>=Input_ProjYears, "", ExcelFunctions.IF(C48=12, B48+1, B48));
    }

    public double cell_C49() {
        return ExcelFunctions.IF(C48=12, 1, C48+1);
    }

    public double cell_D49() {
        return Input_IssueAge+(B49-1);
    }

    public double cell_F49() {
        return ExcelFunctions.IF(B49="", 0, J48);
    }

    public double cell_G49() {
        return VLOOKUP(A49, Premium!48:1095, 12, TRUE);
    }

    public double cell_H49() {
        return ExcelFunctions.IF(B49="", 0, (VLOOKUP(D49, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F49+G49))) + Input_AdminFee);
    }

    public double cell_I49() {
        return ExcelFunctions.IF(B49="", 0, (F49+G49-H49) * Calc_MthlyRate_Cur);
    }

    public double cell_J49() {
        return ExcelFunctions.IF(B49="", 0, Math.max(0, F49+G49-H49+I49));
    }

    public double cell_K49() {
        return ExcelFunctions.IF(B49="", 0, VLOOKUP(D49, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L49() {
        return ExcelFunctions.IF(B49="", 0, Math.max(0, J49-K49));
    }

    public double cell_M49() {
        return ExcelFunctions.IF(B49="", 0, Math.max(Input_FaceAmt, J49 * VLOOKUP(D49, Table_Rates, 4, FALSE)));
    }

    public double cell_O49() {
        return ACV;
    }

    public double cell_P49() {
        return Premium!L49;
    }

    public double cell_Q49() {
        return ExcelFunctions.IF(B49="", 0, ((VLOOKUP(D49, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O49+P49))) + Input_AdminFee);
    }

    public double cell_R49() {
        return ExcelFunctions.IF(B49="", 0, (O49+P49-Q49) * Calc_MthlyRate_Guar);
    }

    public double cell_S49() {
        return ExcelFunctions.IF(B49="", 0, Math.max(0, O49+P49-Q49+R49));
    }

    public double cell_A50() {
        return EDATE(A49, 1);
    }

    public double cell_B50() {
        return ExcelFunctions.IF(B49>=Input_ProjYears, "", ExcelFunctions.IF(C49=12, B49+1, B49));
    }

    public double cell_C50() {
        return ExcelFunctions.IF(C49=12, 1, C49+1);
    }

    public double cell_D50() {
        return Input_IssueAge+(B50-1);
    }

    public double cell_F50() {
        return ExcelFunctions.IF(B50="", 0, J49);
    }

    public double cell_G50() {
        return VLOOKUP(A50, Premium!49:1096, 12, TRUE);
    }

    public double cell_H50() {
        return ExcelFunctions.IF(B50="", 0, (VLOOKUP(D50, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F50+G50))) + Input_AdminFee);
    }

    public double cell_I50() {
        return ExcelFunctions.IF(B50="", 0, (F50+G50-H50) * Calc_MthlyRate_Cur);
    }

    public double cell_J50() {
        return ExcelFunctions.IF(B50="", 0, Math.max(0, F50+G50-H50+I50));
    }

    public double cell_K50() {
        return ExcelFunctions.IF(B50="", 0, VLOOKUP(D50, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L50() {
        return ExcelFunctions.IF(B50="", 0, Math.max(0, J50-K50));
    }

    public double cell_M50() {
        return ExcelFunctions.IF(B50="", 0, Math.max(Input_FaceAmt, J50 * VLOOKUP(D50, Table_Rates, 4, FALSE)));
    }

    public double cell_O50() {
        return ACV;
    }

    public double cell_P50() {
        return Premium!L50;
    }

    public double cell_Q50() {
        return ExcelFunctions.IF(B50="", 0, ((VLOOKUP(D50, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O50+P50))) + Input_AdminFee);
    }

    public double cell_R50() {
        return ExcelFunctions.IF(B50="", 0, (O50+P50-Q50) * Calc_MthlyRate_Guar);
    }

    public double cell_S50() {
        return ExcelFunctions.IF(B50="", 0, Math.max(0, O50+P50-Q50+R50));
    }

    public double cell_A51() {
        return EDATE(A50, 1);
    }

    public double cell_B51() {
        return ExcelFunctions.IF(B50>=Input_ProjYears, "", ExcelFunctions.IF(C50=12, B50+1, B50));
    }

    public double cell_C51() {
        return ExcelFunctions.IF(C50=12, 1, C50+1);
    }

    public double cell_D51() {
        return Input_IssueAge+(B51-1);
    }

    public double cell_F51() {
        return ExcelFunctions.IF(B51="", 0, J50);
    }

    public double cell_G51() {
        return VLOOKUP(A51, Premium!50:1097, 12, TRUE);
    }

    public double cell_H51() {
        return ExcelFunctions.IF(B51="", 0, (VLOOKUP(D51, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F51+G51))) + Input_AdminFee);
    }

    public double cell_I51() {
        return ExcelFunctions.IF(B51="", 0, (F51+G51-H51) * Calc_MthlyRate_Cur);
    }

    public double cell_J51() {
        return ExcelFunctions.IF(B51="", 0, Math.max(0, F51+G51-H51+I51));
    }

    public double cell_K51() {
        return ExcelFunctions.IF(B51="", 0, VLOOKUP(D51, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L51() {
        return ExcelFunctions.IF(B51="", 0, Math.max(0, J51-K51));
    }

    public double cell_M51() {
        return ExcelFunctions.IF(B51="", 0, Math.max(Input_FaceAmt, J51 * VLOOKUP(D51, Table_Rates, 4, FALSE)));
    }

    public double cell_O51() {
        return ACV;
    }

    public double cell_P51() {
        return Premium!L51;
    }

    public double cell_Q51() {
        return ExcelFunctions.IF(B51="", 0, ((VLOOKUP(D51, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O51+P51))) + Input_AdminFee);
    }

    public double cell_R51() {
        return ExcelFunctions.IF(B51="", 0, (O51+P51-Q51) * Calc_MthlyRate_Guar);
    }

    public double cell_S51() {
        return ExcelFunctions.IF(B51="", 0, Math.max(0, O51+P51-Q51+R51));
    }

    public double cell_A52() {
        return EDATE(A51, 1);
    }

    public double cell_B52() {
        return ExcelFunctions.IF(B51>=Input_ProjYears, "", ExcelFunctions.IF(C51=12, B51+1, B51));
    }

    public double cell_C52() {
        return ExcelFunctions.IF(C51=12, 1, C51+1);
    }

    public double cell_D52() {
        return Input_IssueAge+(B52-1);
    }

    public double cell_F52() {
        return ExcelFunctions.IF(B52="", 0, J51);
    }

    public double cell_G52() {
        return VLOOKUP(A52, Premium!51:1098, 12, TRUE);
    }

    public double cell_H52() {
        return ExcelFunctions.IF(B52="", 0, (VLOOKUP(D52, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F52+G52))) + Input_AdminFee);
    }

    public double cell_I52() {
        return ExcelFunctions.IF(B52="", 0, (F52+G52-H52) * Calc_MthlyRate_Cur);
    }

    public double cell_J52() {
        return ExcelFunctions.IF(B52="", 0, Math.max(0, F52+G52-H52+I52));
    }

    public double cell_K52() {
        return ExcelFunctions.IF(B52="", 0, VLOOKUP(D52, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L52() {
        return ExcelFunctions.IF(B52="", 0, Math.max(0, J52-K52));
    }

    public double cell_M52() {
        return ExcelFunctions.IF(B52="", 0, Math.max(Input_FaceAmt, J52 * VLOOKUP(D52, Table_Rates, 4, FALSE)));
    }

    public double cell_O52() {
        return ACV;
    }

    public double cell_P52() {
        return Premium!L52;
    }

    public double cell_Q52() {
        return ExcelFunctions.IF(B52="", 0, ((VLOOKUP(D52, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O52+P52))) + Input_AdminFee);
    }

    public double cell_R52() {
        return ExcelFunctions.IF(B52="", 0, (O52+P52-Q52) * Calc_MthlyRate_Guar);
    }

    public double cell_S52() {
        return ExcelFunctions.IF(B52="", 0, Math.max(0, O52+P52-Q52+R52));
    }

    public double cell_A53() {
        return EDATE(A52, 1);
    }

    public double cell_B53() {
        return ExcelFunctions.IF(B52>=Input_ProjYears, "", ExcelFunctions.IF(C52=12, B52+1, B52));
    }

    public double cell_C53() {
        return ExcelFunctions.IF(C52=12, 1, C52+1);
    }

    public double cell_D53() {
        return Input_IssueAge+(B53-1);
    }

    public double cell_F53() {
        return ExcelFunctions.IF(B53="", 0, J52);
    }

    public double cell_G53() {
        return VLOOKUP(A53, Premium!52:1099, 12, TRUE);
    }

    public double cell_H53() {
        return ExcelFunctions.IF(B53="", 0, (VLOOKUP(D53, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F53+G53))) + Input_AdminFee);
    }

    public double cell_I53() {
        return ExcelFunctions.IF(B53="", 0, (F53+G53-H53) * Calc_MthlyRate_Cur);
    }

    public double cell_J53() {
        return ExcelFunctions.IF(B53="", 0, Math.max(0, F53+G53-H53+I53));
    }

    public double cell_K53() {
        return ExcelFunctions.IF(B53="", 0, VLOOKUP(D53, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L53() {
        return ExcelFunctions.IF(B53="", 0, Math.max(0, J53-K53));
    }

    public double cell_M53() {
        return ExcelFunctions.IF(B53="", 0, Math.max(Input_FaceAmt, J53 * VLOOKUP(D53, Table_Rates, 4, FALSE)));
    }

    public double cell_O53() {
        return ACV;
    }

    public double cell_P53() {
        return Premium!L53;
    }

    public double cell_Q53() {
        return ExcelFunctions.IF(B53="", 0, ((VLOOKUP(D53, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O53+P53))) + Input_AdminFee);
    }

    public double cell_R53() {
        return ExcelFunctions.IF(B53="", 0, (O53+P53-Q53) * Calc_MthlyRate_Guar);
    }

    public double cell_S53() {
        return ExcelFunctions.IF(B53="", 0, Math.max(0, O53+P53-Q53+R53));
    }

    public double cell_A54() {
        return EDATE(A53, 1);
    }

    public double cell_B54() {
        return ExcelFunctions.IF(B53>=Input_ProjYears, "", ExcelFunctions.IF(C53=12, B53+1, B53));
    }

    public double cell_C54() {
        return ExcelFunctions.IF(C53=12, 1, C53+1);
    }

    public double cell_D54() {
        return Input_IssueAge+(B54-1);
    }

    public double cell_F54() {
        return ExcelFunctions.IF(B54="", 0, J53);
    }

    public double cell_G54() {
        return VLOOKUP(A54, Premium!53:1100, 12, TRUE);
    }

    public double cell_H54() {
        return ExcelFunctions.IF(B54="", 0, (VLOOKUP(D54, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F54+G54))) + Input_AdminFee);
    }

    public double cell_I54() {
        return ExcelFunctions.IF(B54="", 0, (F54+G54-H54) * Calc_MthlyRate_Cur);
    }

    public double cell_J54() {
        return ExcelFunctions.IF(B54="", 0, Math.max(0, F54+G54-H54+I54));
    }

    public double cell_K54() {
        return ExcelFunctions.IF(B54="", 0, VLOOKUP(D54, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L54() {
        return ExcelFunctions.IF(B54="", 0, Math.max(0, J54-K54));
    }

    public double cell_M54() {
        return ExcelFunctions.IF(B54="", 0, Math.max(Input_FaceAmt, J54 * VLOOKUP(D54, Table_Rates, 4, FALSE)));
    }

    public double cell_O54() {
        return ACV;
    }

    public double cell_P54() {
        return Premium!L54;
    }

    public double cell_Q54() {
        return ExcelFunctions.IF(B54="", 0, ((VLOOKUP(D54, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O54+P54))) + Input_AdminFee);
    }

    public double cell_R54() {
        return ExcelFunctions.IF(B54="", 0, (O54+P54-Q54) * Calc_MthlyRate_Guar);
    }

    public double cell_S54() {
        return ExcelFunctions.IF(B54="", 0, Math.max(0, O54+P54-Q54+R54));
    }

    public double cell_A55() {
        return EDATE(A54, 1);
    }

    public double cell_B55() {
        return ExcelFunctions.IF(B54>=Input_ProjYears, "", ExcelFunctions.IF(C54=12, B54+1, B54));
    }

    public double cell_C55() {
        return ExcelFunctions.IF(C54=12, 1, C54+1);
    }

    public double cell_D55() {
        return Input_IssueAge+(B55-1);
    }

    public double cell_F55() {
        return ExcelFunctions.IF(B55="", 0, J54);
    }

    public double cell_G55() {
        return VLOOKUP(A55, Premium!54:1101, 12, TRUE);
    }

    public double cell_H55() {
        return ExcelFunctions.IF(B55="", 0, (VLOOKUP(D55, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F55+G55))) + Input_AdminFee);
    }

    public double cell_I55() {
        return ExcelFunctions.IF(B55="", 0, (F55+G55-H55) * Calc_MthlyRate_Cur);
    }

    public double cell_J55() {
        return ExcelFunctions.IF(B55="", 0, Math.max(0, F55+G55-H55+I55));
    }

    public double cell_K55() {
        return ExcelFunctions.IF(B55="", 0, VLOOKUP(D55, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L55() {
        return ExcelFunctions.IF(B55="", 0, Math.max(0, J55-K55));
    }

    public double cell_M55() {
        return ExcelFunctions.IF(B55="", 0, Math.max(Input_FaceAmt, J55 * VLOOKUP(D55, Table_Rates, 4, FALSE)));
    }

    public double cell_O55() {
        return ACV;
    }

    public double cell_P55() {
        return Premium!L55;
    }

    public double cell_Q55() {
        return ExcelFunctions.IF(B55="", 0, ((VLOOKUP(D55, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O55+P55))) + Input_AdminFee);
    }

    public double cell_R55() {
        return ExcelFunctions.IF(B55="", 0, (O55+P55-Q55) * Calc_MthlyRate_Guar);
    }

    public double cell_S55() {
        return ExcelFunctions.IF(B55="", 0, Math.max(0, O55+P55-Q55+R55));
    }

    public double cell_A56() {
        return EDATE(A55, 1);
    }

    public double cell_B56() {
        return ExcelFunctions.IF(B55>=Input_ProjYears, "", ExcelFunctions.IF(C55=12, B55+1, B55));
    }

    public double cell_C56() {
        return ExcelFunctions.IF(C55=12, 1, C55+1);
    }

    public double cell_D56() {
        return Input_IssueAge+(B56-1);
    }

    public double cell_F56() {
        return ExcelFunctions.IF(B56="", 0, J55);
    }

    public double cell_G56() {
        return VLOOKUP(A56, Premium!55:1102, 12, TRUE);
    }

    public double cell_H56() {
        return ExcelFunctions.IF(B56="", 0, (VLOOKUP(D56, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F56+G56))) + Input_AdminFee);
    }

    public double cell_I56() {
        return ExcelFunctions.IF(B56="", 0, (F56+G56-H56) * Calc_MthlyRate_Cur);
    }

    public double cell_J56() {
        return ExcelFunctions.IF(B56="", 0, Math.max(0, F56+G56-H56+I56));
    }

    public double cell_K56() {
        return ExcelFunctions.IF(B56="", 0, VLOOKUP(D56, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L56() {
        return ExcelFunctions.IF(B56="", 0, Math.max(0, J56-K56));
    }

    public double cell_M56() {
        return ExcelFunctions.IF(B56="", 0, Math.max(Input_FaceAmt, J56 * VLOOKUP(D56, Table_Rates, 4, FALSE)));
    }

    public double cell_O56() {
        return ACV;
    }

    public double cell_P56() {
        return Premium!L56;
    }

    public double cell_Q56() {
        return ExcelFunctions.IF(B56="", 0, ((VLOOKUP(D56, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O56+P56))) + Input_AdminFee);
    }

    public double cell_R56() {
        return ExcelFunctions.IF(B56="", 0, (O56+P56-Q56) * Calc_MthlyRate_Guar);
    }

    public double cell_S56() {
        return ExcelFunctions.IF(B56="", 0, Math.max(0, O56+P56-Q56+R56));
    }

    public double cell_A57() {
        return EDATE(A56, 1);
    }

    public double cell_B57() {
        return ExcelFunctions.IF(B56>=Input_ProjYears, "", ExcelFunctions.IF(C56=12, B56+1, B56));
    }

    public double cell_C57() {
        return ExcelFunctions.IF(C56=12, 1, C56+1);
    }

    public double cell_D57() {
        return Input_IssueAge+(B57-1);
    }

    public double cell_F57() {
        return ExcelFunctions.IF(B57="", 0, J56);
    }

    public double cell_G57() {
        return VLOOKUP(A57, Premium!56:1103, 12, TRUE);
    }

    public double cell_H57() {
        return ExcelFunctions.IF(B57="", 0, (VLOOKUP(D57, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F57+G57))) + Input_AdminFee);
    }

    public double cell_I57() {
        return ExcelFunctions.IF(B57="", 0, (F57+G57-H57) * Calc_MthlyRate_Cur);
    }

    public double cell_J57() {
        return ExcelFunctions.IF(B57="", 0, Math.max(0, F57+G57-H57+I57));
    }

    public double cell_K57() {
        return ExcelFunctions.IF(B57="", 0, VLOOKUP(D57, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L57() {
        return ExcelFunctions.IF(B57="", 0, Math.max(0, J57-K57));
    }

    public double cell_M57() {
        return ExcelFunctions.IF(B57="", 0, Math.max(Input_FaceAmt, J57 * VLOOKUP(D57, Table_Rates, 4, FALSE)));
    }

    public double cell_O57() {
        return ACV;
    }

    public double cell_P57() {
        return Premium!L57;
    }

    public double cell_Q57() {
        return ExcelFunctions.IF(B57="", 0, ((VLOOKUP(D57, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O57+P57))) + Input_AdminFee);
    }

    public double cell_R57() {
        return ExcelFunctions.IF(B57="", 0, (O57+P57-Q57) * Calc_MthlyRate_Guar);
    }

    public double cell_S57() {
        return ExcelFunctions.IF(B57="", 0, Math.max(0, O57+P57-Q57+R57));
    }

    public double cell_A58() {
        return EDATE(A57, 1);
    }

    public double cell_B58() {
        return ExcelFunctions.IF(B57>=Input_ProjYears, "", ExcelFunctions.IF(C57=12, B57+1, B57));
    }

    public double cell_C58() {
        return ExcelFunctions.IF(C57=12, 1, C57+1);
    }

    public double cell_D58() {
        return Input_IssueAge+(B58-1);
    }

    public double cell_F58() {
        return ExcelFunctions.IF(B58="", 0, J57);
    }

    public double cell_G58() {
        return VLOOKUP(A58, Premium!57:1104, 12, TRUE);
    }

    public double cell_H58() {
        return ExcelFunctions.IF(B58="", 0, (VLOOKUP(D58, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F58+G58))) + Input_AdminFee);
    }

    public double cell_I58() {
        return ExcelFunctions.IF(B58="", 0, (F58+G58-H58) * Calc_MthlyRate_Cur);
    }

    public double cell_J58() {
        return ExcelFunctions.IF(B58="", 0, Math.max(0, F58+G58-H58+I58));
    }

    public double cell_K58() {
        return ExcelFunctions.IF(B58="", 0, VLOOKUP(D58, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L58() {
        return ExcelFunctions.IF(B58="", 0, Math.max(0, J58-K58));
    }

    public double cell_M58() {
        return ExcelFunctions.IF(B58="", 0, Math.max(Input_FaceAmt, J58 * VLOOKUP(D58, Table_Rates, 4, FALSE)));
    }

    public double cell_O58() {
        return ACV;
    }

    public double cell_P58() {
        return Premium!L58;
    }

    public double cell_Q58() {
        return ExcelFunctions.IF(B58="", 0, ((VLOOKUP(D58, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O58+P58))) + Input_AdminFee);
    }

    public double cell_R58() {
        return ExcelFunctions.IF(B58="", 0, (O58+P58-Q58) * Calc_MthlyRate_Guar);
    }

    public double cell_S58() {
        return ExcelFunctions.IF(B58="", 0, Math.max(0, O58+P58-Q58+R58));
    }

    public double cell_A59() {
        return EDATE(A58, 1);
    }

    public double cell_B59() {
        return ExcelFunctions.IF(B58>=Input_ProjYears, "", ExcelFunctions.IF(C58=12, B58+1, B58));
    }

    public double cell_C59() {
        return ExcelFunctions.IF(C58=12, 1, C58+1);
    }

    public double cell_D59() {
        return Input_IssueAge+(B59-1);
    }

    public double cell_F59() {
        return ExcelFunctions.IF(B59="", 0, J58);
    }

    public double cell_G59() {
        return VLOOKUP(A59, Premium!58:1105, 12, TRUE);
    }

    public double cell_H59() {
        return ExcelFunctions.IF(B59="", 0, (VLOOKUP(D59, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F59+G59))) + Input_AdminFee);
    }

    public double cell_I59() {
        return ExcelFunctions.IF(B59="", 0, (F59+G59-H59) * Calc_MthlyRate_Cur);
    }

    public double cell_J59() {
        return ExcelFunctions.IF(B59="", 0, Math.max(0, F59+G59-H59+I59));
    }

    public double cell_K59() {
        return ExcelFunctions.IF(B59="", 0, VLOOKUP(D59, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L59() {
        return ExcelFunctions.IF(B59="", 0, Math.max(0, J59-K59));
    }

    public double cell_M59() {
        return ExcelFunctions.IF(B59="", 0, Math.max(Input_FaceAmt, J59 * VLOOKUP(D59, Table_Rates, 4, FALSE)));
    }

    public double cell_O59() {
        return ACV;
    }

    public double cell_P59() {
        return Premium!L59;
    }

    public double cell_Q59() {
        return ExcelFunctions.IF(B59="", 0, ((VLOOKUP(D59, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O59+P59))) + Input_AdminFee);
    }

    public double cell_R59() {
        return ExcelFunctions.IF(B59="", 0, (O59+P59-Q59) * Calc_MthlyRate_Guar);
    }

    public double cell_S59() {
        return ExcelFunctions.IF(B59="", 0, Math.max(0, O59+P59-Q59+R59));
    }

    public double cell_A60() {
        return EDATE(A59, 1);
    }

    public double cell_B60() {
        return ExcelFunctions.IF(B59>=Input_ProjYears, "", ExcelFunctions.IF(C59=12, B59+1, B59));
    }

    public double cell_C60() {
        return ExcelFunctions.IF(C59=12, 1, C59+1);
    }

    public double cell_D60() {
        return Input_IssueAge+(B60-1);
    }

    public double cell_F60() {
        return ExcelFunctions.IF(B60="", 0, J59);
    }

    public double cell_G60() {
        return VLOOKUP(A60, Premium!59:1106, 12, TRUE);
    }

    public double cell_H60() {
        return ExcelFunctions.IF(B60="", 0, (VLOOKUP(D60, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F60+G60))) + Input_AdminFee);
    }

    public double cell_I60() {
        return ExcelFunctions.IF(B60="", 0, (F60+G60-H60) * Calc_MthlyRate_Cur);
    }

    public double cell_J60() {
        return ExcelFunctions.IF(B60="", 0, Math.max(0, F60+G60-H60+I60));
    }

    public double cell_K60() {
        return ExcelFunctions.IF(B60="", 0, VLOOKUP(D60, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L60() {
        return ExcelFunctions.IF(B60="", 0, Math.max(0, J60-K60));
    }

    public double cell_M60() {
        return ExcelFunctions.IF(B60="", 0, Math.max(Input_FaceAmt, J60 * VLOOKUP(D60, Table_Rates, 4, FALSE)));
    }

    public double cell_O60() {
        return ACV;
    }

    public double cell_P60() {
        return Premium!L60;
    }

    public double cell_Q60() {
        return ExcelFunctions.IF(B60="", 0, ((VLOOKUP(D60, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O60+P60))) + Input_AdminFee);
    }

    public double cell_R60() {
        return ExcelFunctions.IF(B60="", 0, (O60+P60-Q60) * Calc_MthlyRate_Guar);
    }

    public double cell_S60() {
        return ExcelFunctions.IF(B60="", 0, Math.max(0, O60+P60-Q60+R60));
    }

    public double cell_A61() {
        return EDATE(A60, 1);
    }

    public double cell_B61() {
        return ExcelFunctions.IF(B60>=Input_ProjYears, "", ExcelFunctions.IF(C60=12, B60+1, B60));
    }

    public double cell_C61() {
        return ExcelFunctions.IF(C60=12, 1, C60+1);
    }

    public double cell_D61() {
        return Input_IssueAge+(B61-1);
    }

    public double cell_F61() {
        return ExcelFunctions.IF(B61="", 0, J60);
    }

    public double cell_G61() {
        return VLOOKUP(A61, Premium!60:1107, 12, TRUE);
    }

    public double cell_H61() {
        return ExcelFunctions.IF(B61="", 0, (VLOOKUP(D61, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F61+G61))) + Input_AdminFee);
    }

    public double cell_I61() {
        return ExcelFunctions.IF(B61="", 0, (F61+G61-H61) * Calc_MthlyRate_Cur);
    }

    public double cell_J61() {
        return ExcelFunctions.IF(B61="", 0, Math.max(0, F61+G61-H61+I61));
    }

    public double cell_K61() {
        return ExcelFunctions.IF(B61="", 0, VLOOKUP(D61, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L61() {
        return ExcelFunctions.IF(B61="", 0, Math.max(0, J61-K61));
    }

    public double cell_M61() {
        return ExcelFunctions.IF(B61="", 0, Math.max(Input_FaceAmt, J61 * VLOOKUP(D61, Table_Rates, 4, FALSE)));
    }

    public double cell_O61() {
        return ACV;
    }

    public double cell_P61() {
        return Premium!L61;
    }

    public double cell_Q61() {
        return ExcelFunctions.IF(B61="", 0, ((VLOOKUP(D61, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O61+P61))) + Input_AdminFee);
    }

    public double cell_R61() {
        return ExcelFunctions.IF(B61="", 0, (O61+P61-Q61) * Calc_MthlyRate_Guar);
    }

    public double cell_S61() {
        return ExcelFunctions.IF(B61="", 0, Math.max(0, O61+P61-Q61+R61));
    }

    public double cell_A62() {
        return EDATE(A61, 1);
    }

    public double cell_B62() {
        return ExcelFunctions.IF(B61>=Input_ProjYears, "", ExcelFunctions.IF(C61=12, B61+1, B61));
    }

    public double cell_C62() {
        return ExcelFunctions.IF(C61=12, 1, C61+1);
    }

    public double cell_D62() {
        return Input_IssueAge+(B62-1);
    }

    public double cell_F62() {
        return ExcelFunctions.IF(B62="", 0, J61);
    }

    public double cell_G62() {
        return VLOOKUP(A62, Premium!61:1108, 12, TRUE);
    }

    public double cell_H62() {
        return ExcelFunctions.IF(B62="", 0, (VLOOKUP(D62, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F62+G62))) + Input_AdminFee);
    }

    public double cell_I62() {
        return ExcelFunctions.IF(B62="", 0, (F62+G62-H62) * Calc_MthlyRate_Cur);
    }

    public double cell_J62() {
        return ExcelFunctions.IF(B62="", 0, Math.max(0, F62+G62-H62+I62));
    }

    public double cell_K62() {
        return ExcelFunctions.IF(B62="", 0, VLOOKUP(D62, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L62() {
        return ExcelFunctions.IF(B62="", 0, Math.max(0, J62-K62));
    }

    public double cell_M62() {
        return ExcelFunctions.IF(B62="", 0, Math.max(Input_FaceAmt, J62 * VLOOKUP(D62, Table_Rates, 4, FALSE)));
    }

    public double cell_O62() {
        return ACV;
    }

    public double cell_P62() {
        return Premium!L62;
    }

    public double cell_Q62() {
        return ExcelFunctions.IF(B62="", 0, ((VLOOKUP(D62, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O62+P62))) + Input_AdminFee);
    }

    public double cell_R62() {
        return ExcelFunctions.IF(B62="", 0, (O62+P62-Q62) * Calc_MthlyRate_Guar);
    }

    public double cell_S62() {
        return ExcelFunctions.IF(B62="", 0, Math.max(0, O62+P62-Q62+R62));
    }

    public double cell_A63() {
        return EDATE(A62, 1);
    }

    public double cell_B63() {
        return ExcelFunctions.IF(B62>=Input_ProjYears, "", ExcelFunctions.IF(C62=12, B62+1, B62));
    }

    public double cell_C63() {
        return ExcelFunctions.IF(C62=12, 1, C62+1);
    }

    public double cell_D63() {
        return Input_IssueAge+(B63-1);
    }

    public double cell_F63() {
        return ExcelFunctions.IF(B63="", 0, J62);
    }

    public double cell_G63() {
        return VLOOKUP(A63, Premium!62:1109, 12, TRUE);
    }

    public double cell_H63() {
        return ExcelFunctions.IF(B63="", 0, (VLOOKUP(D63, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F63+G63))) + Input_AdminFee);
    }

    public double cell_I63() {
        return ExcelFunctions.IF(B63="", 0, (F63+G63-H63) * Calc_MthlyRate_Cur);
    }

    public double cell_J63() {
        return ExcelFunctions.IF(B63="", 0, Math.max(0, F63+G63-H63+I63));
    }

    public double cell_K63() {
        return ExcelFunctions.IF(B63="", 0, VLOOKUP(D63, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L63() {
        return ExcelFunctions.IF(B63="", 0, Math.max(0, J63-K63));
    }

    public double cell_M63() {
        return ExcelFunctions.IF(B63="", 0, Math.max(Input_FaceAmt, J63 * VLOOKUP(D63, Table_Rates, 4, FALSE)));
    }

    public double cell_O63() {
        return ACV;
    }

    public double cell_P63() {
        return Premium!L63;
    }

    public double cell_Q63() {
        return ExcelFunctions.IF(B63="", 0, ((VLOOKUP(D63, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O63+P63))) + Input_AdminFee);
    }

    public double cell_R63() {
        return ExcelFunctions.IF(B63="", 0, (O63+P63-Q63) * Calc_MthlyRate_Guar);
    }

    public double cell_S63() {
        return ExcelFunctions.IF(B63="", 0, Math.max(0, O63+P63-Q63+R63));
    }

    public double cell_A64() {
        return EDATE(A63, 1);
    }

    public double cell_B64() {
        return ExcelFunctions.IF(B63>=Input_ProjYears, "", ExcelFunctions.IF(C63=12, B63+1, B63));
    }

    public double cell_C64() {
        return ExcelFunctions.IF(C63=12, 1, C63+1);
    }

    public double cell_D64() {
        return Input_IssueAge+(B64-1);
    }

    public double cell_F64() {
        return ExcelFunctions.IF(B64="", 0, J63);
    }

    public double cell_G64() {
        return VLOOKUP(A64, Premium!63:1110, 12, TRUE);
    }

    public double cell_H64() {
        return ExcelFunctions.IF(B64="", 0, (VLOOKUP(D64, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F64+G64))) + Input_AdminFee);
    }

    public double cell_I64() {
        return ExcelFunctions.IF(B64="", 0, (F64+G64-H64) * Calc_MthlyRate_Cur);
    }

    public double cell_J64() {
        return ExcelFunctions.IF(B64="", 0, Math.max(0, F64+G64-H64+I64));
    }

    public double cell_K64() {
        return ExcelFunctions.IF(B64="", 0, VLOOKUP(D64, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L64() {
        return ExcelFunctions.IF(B64="", 0, Math.max(0, J64-K64));
    }

    public double cell_M64() {
        return ExcelFunctions.IF(B64="", 0, Math.max(Input_FaceAmt, J64 * VLOOKUP(D64, Table_Rates, 4, FALSE)));
    }

    public double cell_O64() {
        return ACV;
    }

    public double cell_P64() {
        return Premium!L64;
    }

    public double cell_Q64() {
        return ExcelFunctions.IF(B64="", 0, ((VLOOKUP(D64, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O64+P64))) + Input_AdminFee);
    }

    public double cell_R64() {
        return ExcelFunctions.IF(B64="", 0, (O64+P64-Q64) * Calc_MthlyRate_Guar);
    }

    public double cell_S64() {
        return ExcelFunctions.IF(B64="", 0, Math.max(0, O64+P64-Q64+R64));
    }

    public double cell_A65() {
        return EDATE(A64, 1);
    }

    public double cell_B65() {
        return ExcelFunctions.IF(B64>=Input_ProjYears, "", ExcelFunctions.IF(C64=12, B64+1, B64));
    }

    public double cell_C65() {
        return ExcelFunctions.IF(C64=12, 1, C64+1);
    }

    public double cell_D65() {
        return Input_IssueAge+(B65-1);
    }

    public double cell_F65() {
        return ExcelFunctions.IF(B65="", 0, J64);
    }

    public double cell_G65() {
        return VLOOKUP(A65, Premium!64:1111, 12, TRUE);
    }

    public double cell_H65() {
        return ExcelFunctions.IF(B65="", 0, (VLOOKUP(D65, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F65+G65))) + Input_AdminFee);
    }

    public double cell_I65() {
        return ExcelFunctions.IF(B65="", 0, (F65+G65-H65) * Calc_MthlyRate_Cur);
    }

    public double cell_J65() {
        return ExcelFunctions.IF(B65="", 0, Math.max(0, F65+G65-H65+I65));
    }

    public double cell_K65() {
        return ExcelFunctions.IF(B65="", 0, VLOOKUP(D65, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L65() {
        return ExcelFunctions.IF(B65="", 0, Math.max(0, J65-K65));
    }

    public double cell_M65() {
        return ExcelFunctions.IF(B65="", 0, Math.max(Input_FaceAmt, J65 * VLOOKUP(D65, Table_Rates, 4, FALSE)));
    }

    public double cell_O65() {
        return ACV;
    }

    public double cell_P65() {
        return Premium!L65;
    }

    public double cell_Q65() {
        return ExcelFunctions.IF(B65="", 0, ((VLOOKUP(D65, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O65+P65))) + Input_AdminFee);
    }

    public double cell_R65() {
        return ExcelFunctions.IF(B65="", 0, (O65+P65-Q65) * Calc_MthlyRate_Guar);
    }

    public double cell_S65() {
        return ExcelFunctions.IF(B65="", 0, Math.max(0, O65+P65-Q65+R65));
    }

    public double cell_A66() {
        return EDATE(A65, 1);
    }

    public double cell_B66() {
        return ExcelFunctions.IF(B65>=Input_ProjYears, "", ExcelFunctions.IF(C65=12, B65+1, B65));
    }

    public double cell_C66() {
        return ExcelFunctions.IF(C65=12, 1, C65+1);
    }

    public double cell_D66() {
        return Input_IssueAge+(B66-1);
    }

    public double cell_F66() {
        return ExcelFunctions.IF(B66="", 0, J65);
    }

    public double cell_G66() {
        return VLOOKUP(A66, Premium!65:1112, 12, TRUE);
    }

    public double cell_H66() {
        return ExcelFunctions.IF(B66="", 0, (VLOOKUP(D66, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F66+G66))) + Input_AdminFee);
    }

    public double cell_I66() {
        return ExcelFunctions.IF(B66="", 0, (F66+G66-H66) * Calc_MthlyRate_Cur);
    }

    public double cell_J66() {
        return ExcelFunctions.IF(B66="", 0, Math.max(0, F66+G66-H66+I66));
    }

    public double cell_K66() {
        return ExcelFunctions.IF(B66="", 0, VLOOKUP(D66, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L66() {
        return ExcelFunctions.IF(B66="", 0, Math.max(0, J66-K66));
    }

    public double cell_M66() {
        return ExcelFunctions.IF(B66="", 0, Math.max(Input_FaceAmt, J66 * VLOOKUP(D66, Table_Rates, 4, FALSE)));
    }

    public double cell_O66() {
        return ACV;
    }

    public double cell_P66() {
        return Premium!L66;
    }

    public double cell_Q66() {
        return ExcelFunctions.IF(B66="", 0, ((VLOOKUP(D66, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O66+P66))) + Input_AdminFee);
    }

    public double cell_R66() {
        return ExcelFunctions.IF(B66="", 0, (O66+P66-Q66) * Calc_MthlyRate_Guar);
    }

    public double cell_S66() {
        return ExcelFunctions.IF(B66="", 0, Math.max(0, O66+P66-Q66+R66));
    }

    public double cell_A67() {
        return EDATE(A66, 1);
    }

    public double cell_B67() {
        return ExcelFunctions.IF(B66>=Input_ProjYears, "", ExcelFunctions.IF(C66=12, B66+1, B66));
    }

    public double cell_C67() {
        return ExcelFunctions.IF(C66=12, 1, C66+1);
    }

    public double cell_D67() {
        return Input_IssueAge+(B67-1);
    }

    public double cell_F67() {
        return ExcelFunctions.IF(B67="", 0, J66);
    }

    public double cell_G67() {
        return VLOOKUP(A67, Premium!66:1113, 12, TRUE);
    }

    public double cell_H67() {
        return ExcelFunctions.IF(B67="", 0, (VLOOKUP(D67, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F67+G67))) + Input_AdminFee);
    }

    public double cell_I67() {
        return ExcelFunctions.IF(B67="", 0, (F67+G67-H67) * Calc_MthlyRate_Cur);
    }

    public double cell_J67() {
        return ExcelFunctions.IF(B67="", 0, Math.max(0, F67+G67-H67+I67));
    }

    public double cell_K67() {
        return ExcelFunctions.IF(B67="", 0, VLOOKUP(D67, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L67() {
        return ExcelFunctions.IF(B67="", 0, Math.max(0, J67-K67));
    }

    public double cell_M67() {
        return ExcelFunctions.IF(B67="", 0, Math.max(Input_FaceAmt, J67 * VLOOKUP(D67, Table_Rates, 4, FALSE)));
    }

    public double cell_O67() {
        return ACV;
    }

    public double cell_P67() {
        return Premium!L67;
    }

    public double cell_Q67() {
        return ExcelFunctions.IF(B67="", 0, ((VLOOKUP(D67, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O67+P67))) + Input_AdminFee);
    }

    public double cell_R67() {
        return ExcelFunctions.IF(B67="", 0, (O67+P67-Q67) * Calc_MthlyRate_Guar);
    }

    public double cell_S67() {
        return ExcelFunctions.IF(B67="", 0, Math.max(0, O67+P67-Q67+R67));
    }

    public double cell_A68() {
        return EDATE(A67, 1);
    }

    public double cell_B68() {
        return ExcelFunctions.IF(B67>=Input_ProjYears, "", ExcelFunctions.IF(C67=12, B67+1, B67));
    }

    public double cell_C68() {
        return ExcelFunctions.IF(C67=12, 1, C67+1);
    }

    public double cell_D68() {
        return Input_IssueAge+(B68-1);
    }

    public double cell_F68() {
        return ExcelFunctions.IF(B68="", 0, J67);
    }

    public double cell_G68() {
        return VLOOKUP(A68, Premium!67:1114, 12, TRUE);
    }

    public double cell_H68() {
        return ExcelFunctions.IF(B68="", 0, (VLOOKUP(D68, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F68+G68))) + Input_AdminFee);
    }

    public double cell_I68() {
        return ExcelFunctions.IF(B68="", 0, (F68+G68-H68) * Calc_MthlyRate_Cur);
    }

    public double cell_J68() {
        return ExcelFunctions.IF(B68="", 0, Math.max(0, F68+G68-H68+I68));
    }

    public double cell_K68() {
        return ExcelFunctions.IF(B68="", 0, VLOOKUP(D68, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L68() {
        return ExcelFunctions.IF(B68="", 0, Math.max(0, J68-K68));
    }

    public double cell_M68() {
        return ExcelFunctions.IF(B68="", 0, Math.max(Input_FaceAmt, J68 * VLOOKUP(D68, Table_Rates, 4, FALSE)));
    }

    public double cell_O68() {
        return ACV;
    }

    public double cell_P68() {
        return Premium!L68;
    }

    public double cell_Q68() {
        return ExcelFunctions.IF(B68="", 0, ((VLOOKUP(D68, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O68+P68))) + Input_AdminFee);
    }

    public double cell_R68() {
        return ExcelFunctions.IF(B68="", 0, (O68+P68-Q68) * Calc_MthlyRate_Guar);
    }

    public double cell_S68() {
        return ExcelFunctions.IF(B68="", 0, Math.max(0, O68+P68-Q68+R68));
    }

    public double cell_A69() {
        return EDATE(A68, 1);
    }

    public double cell_B69() {
        return ExcelFunctions.IF(B68>=Input_ProjYears, "", ExcelFunctions.IF(C68=12, B68+1, B68));
    }

    public double cell_C69() {
        return ExcelFunctions.IF(C68=12, 1, C68+1);
    }

    public double cell_D69() {
        return Input_IssueAge+(B69-1);
    }

    public double cell_F69() {
        return ExcelFunctions.IF(B69="", 0, J68);
    }

    public double cell_G69() {
        return VLOOKUP(A69, Premium!68:1115, 12, TRUE);
    }

    public double cell_H69() {
        return ExcelFunctions.IF(B69="", 0, (VLOOKUP(D69, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F69+G69))) + Input_AdminFee);
    }

    public double cell_I69() {
        return ExcelFunctions.IF(B69="", 0, (F69+G69-H69) * Calc_MthlyRate_Cur);
    }

    public double cell_J69() {
        return ExcelFunctions.IF(B69="", 0, Math.max(0, F69+G69-H69+I69));
    }

    public double cell_K69() {
        return ExcelFunctions.IF(B69="", 0, VLOOKUP(D69, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L69() {
        return ExcelFunctions.IF(B69="", 0, Math.max(0, J69-K69));
    }

    public double cell_M69() {
        return ExcelFunctions.IF(B69="", 0, Math.max(Input_FaceAmt, J69 * VLOOKUP(D69, Table_Rates, 4, FALSE)));
    }

    public double cell_O69() {
        return ACV;
    }

    public double cell_P69() {
        return Premium!L69;
    }

    public double cell_Q69() {
        return ExcelFunctions.IF(B69="", 0, ((VLOOKUP(D69, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O69+P69))) + Input_AdminFee);
    }

    public double cell_R69() {
        return ExcelFunctions.IF(B69="", 0, (O69+P69-Q69) * Calc_MthlyRate_Guar);
    }

    public double cell_S69() {
        return ExcelFunctions.IF(B69="", 0, Math.max(0, O69+P69-Q69+R69));
    }

    public double cell_A70() {
        return EDATE(A69, 1);
    }

    public double cell_B70() {
        return ExcelFunctions.IF(B69>=Input_ProjYears, "", ExcelFunctions.IF(C69=12, B69+1, B69));
    }

    public double cell_C70() {
        return ExcelFunctions.IF(C69=12, 1, C69+1);
    }

    public double cell_D70() {
        return Input_IssueAge+(B70-1);
    }

    public double cell_F70() {
        return ExcelFunctions.IF(B70="", 0, J69);
    }

    public double cell_G70() {
        return VLOOKUP(A70, Premium!69:1116, 12, TRUE);
    }

    public double cell_H70() {
        return ExcelFunctions.IF(B70="", 0, (VLOOKUP(D70, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F70+G70))) + Input_AdminFee);
    }

    public double cell_I70() {
        return ExcelFunctions.IF(B70="", 0, (F70+G70-H70) * Calc_MthlyRate_Cur);
    }

    public double cell_J70() {
        return ExcelFunctions.IF(B70="", 0, Math.max(0, F70+G70-H70+I70));
    }

    public double cell_K70() {
        return ExcelFunctions.IF(B70="", 0, VLOOKUP(D70, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L70() {
        return ExcelFunctions.IF(B70="", 0, Math.max(0, J70-K70));
    }

    public double cell_M70() {
        return ExcelFunctions.IF(B70="", 0, Math.max(Input_FaceAmt, J70 * VLOOKUP(D70, Table_Rates, 4, FALSE)));
    }

    public double cell_O70() {
        return ACV;
    }

    public double cell_P70() {
        return Premium!L70;
    }

    public double cell_Q70() {
        return ExcelFunctions.IF(B70="", 0, ((VLOOKUP(D70, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O70+P70))) + Input_AdminFee);
    }

    public double cell_R70() {
        return ExcelFunctions.IF(B70="", 0, (O70+P70-Q70) * Calc_MthlyRate_Guar);
    }

    public double cell_S70() {
        return ExcelFunctions.IF(B70="", 0, Math.max(0, O70+P70-Q70+R70));
    }

    public double cell_A71() {
        return EDATE(A70, 1);
    }

    public double cell_B71() {
        return ExcelFunctions.IF(B70>=Input_ProjYears, "", ExcelFunctions.IF(C70=12, B70+1, B70));
    }

    public double cell_C71() {
        return ExcelFunctions.IF(C70=12, 1, C70+1);
    }

    public double cell_D71() {
        return Input_IssueAge+(B71-1);
    }

    public double cell_F71() {
        return ExcelFunctions.IF(B71="", 0, J70);
    }

    public double cell_G71() {
        return VLOOKUP(A71, Premium!70:1117, 12, TRUE);
    }

    public double cell_H71() {
        return ExcelFunctions.IF(B71="", 0, (VLOOKUP(D71, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F71+G71))) + Input_AdminFee);
    }

    public double cell_I71() {
        return ExcelFunctions.IF(B71="", 0, (F71+G71-H71) * Calc_MthlyRate_Cur);
    }

    public double cell_J71() {
        return ExcelFunctions.IF(B71="", 0, Math.max(0, F71+G71-H71+I71));
    }

    public double cell_K71() {
        return ExcelFunctions.IF(B71="", 0, VLOOKUP(D71, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L71() {
        return ExcelFunctions.IF(B71="", 0, Math.max(0, J71-K71));
    }

    public double cell_M71() {
        return ExcelFunctions.IF(B71="", 0, Math.max(Input_FaceAmt, J71 * VLOOKUP(D71, Table_Rates, 4, FALSE)));
    }

    public double cell_O71() {
        return ACV;
    }

    public double cell_P71() {
        return Premium!L71;
    }

    public double cell_Q71() {
        return ExcelFunctions.IF(B71="", 0, ((VLOOKUP(D71, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O71+P71))) + Input_AdminFee);
    }

    public double cell_R71() {
        return ExcelFunctions.IF(B71="", 0, (O71+P71-Q71) * Calc_MthlyRate_Guar);
    }

    public double cell_S71() {
        return ExcelFunctions.IF(B71="", 0, Math.max(0, O71+P71-Q71+R71));
    }

    public double cell_A72() {
        return EDATE(A71, 1);
    }

    public double cell_B72() {
        return ExcelFunctions.IF(B71>=Input_ProjYears, "", ExcelFunctions.IF(C71=12, B71+1, B71));
    }

    public double cell_C72() {
        return ExcelFunctions.IF(C71=12, 1, C71+1);
    }

    public double cell_D72() {
        return Input_IssueAge+(B72-1);
    }

    public double cell_F72() {
        return ExcelFunctions.IF(B72="", 0, J71);
    }

    public double cell_G72() {
        return VLOOKUP(A72, Premium!71:1118, 12, TRUE);
    }

    public double cell_H72() {
        return ExcelFunctions.IF(B72="", 0, (VLOOKUP(D72, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F72+G72))) + Input_AdminFee);
    }

    public double cell_I72() {
        return ExcelFunctions.IF(B72="", 0, (F72+G72-H72) * Calc_MthlyRate_Cur);
    }

    public double cell_J72() {
        return ExcelFunctions.IF(B72="", 0, Math.max(0, F72+G72-H72+I72));
    }

    public double cell_K72() {
        return ExcelFunctions.IF(B72="", 0, VLOOKUP(D72, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L72() {
        return ExcelFunctions.IF(B72="", 0, Math.max(0, J72-K72));
    }

    public double cell_M72() {
        return ExcelFunctions.IF(B72="", 0, Math.max(Input_FaceAmt, J72 * VLOOKUP(D72, Table_Rates, 4, FALSE)));
    }

    public double cell_O72() {
        return ACV;
    }

    public double cell_P72() {
        return Premium!L72;
    }

    public double cell_Q72() {
        return ExcelFunctions.IF(B72="", 0, ((VLOOKUP(D72, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O72+P72))) + Input_AdminFee);
    }

    public double cell_R72() {
        return ExcelFunctions.IF(B72="", 0, (O72+P72-Q72) * Calc_MthlyRate_Guar);
    }

    public double cell_S72() {
        return ExcelFunctions.IF(B72="", 0, Math.max(0, O72+P72-Q72+R72));
    }

    public double cell_A73() {
        return EDATE(A72, 1);
    }

    public double cell_B73() {
        return ExcelFunctions.IF(B72>=Input_ProjYears, "", ExcelFunctions.IF(C72=12, B72+1, B72));
    }

    public double cell_C73() {
        return ExcelFunctions.IF(C72=12, 1, C72+1);
    }

    public double cell_D73() {
        return Input_IssueAge+(B73-1);
    }

    public double cell_F73() {
        return ExcelFunctions.IF(B73="", 0, J72);
    }

    public double cell_G73() {
        return VLOOKUP(A73, Premium!72:1119, 12, TRUE);
    }

    public double cell_H73() {
        return ExcelFunctions.IF(B73="", 0, (VLOOKUP(D73, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F73+G73))) + Input_AdminFee);
    }

    public double cell_I73() {
        return ExcelFunctions.IF(B73="", 0, (F73+G73-H73) * Calc_MthlyRate_Cur);
    }

    public double cell_J73() {
        return ExcelFunctions.IF(B73="", 0, Math.max(0, F73+G73-H73+I73));
    }

    public double cell_K73() {
        return ExcelFunctions.IF(B73="", 0, VLOOKUP(D73, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L73() {
        return ExcelFunctions.IF(B73="", 0, Math.max(0, J73-K73));
    }

    public double cell_M73() {
        return ExcelFunctions.IF(B73="", 0, Math.max(Input_FaceAmt, J73 * VLOOKUP(D73, Table_Rates, 4, FALSE)));
    }

    public double cell_O73() {
        return ACV;
    }

    public double cell_P73() {
        return Premium!L73;
    }

    public double cell_Q73() {
        return ExcelFunctions.IF(B73="", 0, ((VLOOKUP(D73, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O73+P73))) + Input_AdminFee);
    }

    public double cell_R73() {
        return ExcelFunctions.IF(B73="", 0, (O73+P73-Q73) * Calc_MthlyRate_Guar);
    }

    public double cell_S73() {
        return ExcelFunctions.IF(B73="", 0, Math.max(0, O73+P73-Q73+R73));
    }

    public double cell_A74() {
        return EDATE(A73, 1);
    }

    public double cell_B74() {
        return ExcelFunctions.IF(B73>=Input_ProjYears, "", ExcelFunctions.IF(C73=12, B73+1, B73));
    }

    public double cell_C74() {
        return ExcelFunctions.IF(C73=12, 1, C73+1);
    }

    public double cell_D74() {
        return Input_IssueAge+(B74-1);
    }

    public double cell_F74() {
        return ExcelFunctions.IF(B74="", 0, J73);
    }

    public double cell_G74() {
        return VLOOKUP(A74, Premium!73:1120, 12, TRUE);
    }

    public double cell_H74() {
        return ExcelFunctions.IF(B74="", 0, (VLOOKUP(D74, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F74+G74))) + Input_AdminFee);
    }

    public double cell_I74() {
        return ExcelFunctions.IF(B74="", 0, (F74+G74-H74) * Calc_MthlyRate_Cur);
    }

    public double cell_J74() {
        return ExcelFunctions.IF(B74="", 0, Math.max(0, F74+G74-H74+I74));
    }

    public double cell_K74() {
        return ExcelFunctions.IF(B74="", 0, VLOOKUP(D74, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L74() {
        return ExcelFunctions.IF(B74="", 0, Math.max(0, J74-K74));
    }

    public double cell_M74() {
        return ExcelFunctions.IF(B74="", 0, Math.max(Input_FaceAmt, J74 * VLOOKUP(D74, Table_Rates, 4, FALSE)));
    }

    public double cell_O74() {
        return ACV;
    }

    public double cell_P74() {
        return Premium!L74;
    }

    public double cell_Q74() {
        return ExcelFunctions.IF(B74="", 0, ((VLOOKUP(D74, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O74+P74))) + Input_AdminFee);
    }

    public double cell_R74() {
        return ExcelFunctions.IF(B74="", 0, (O74+P74-Q74) * Calc_MthlyRate_Guar);
    }

    public double cell_S74() {
        return ExcelFunctions.IF(B74="", 0, Math.max(0, O74+P74-Q74+R74));
    }

    public double cell_A75() {
        return EDATE(A74, 1);
    }

    public double cell_B75() {
        return ExcelFunctions.IF(B74>=Input_ProjYears, "", ExcelFunctions.IF(C74=12, B74+1, B74));
    }

    public double cell_C75() {
        return ExcelFunctions.IF(C74=12, 1, C74+1);
    }

    public double cell_D75() {
        return Input_IssueAge+(B75-1);
    }

    public double cell_F75() {
        return ExcelFunctions.IF(B75="", 0, J74);
    }

    public double cell_G75() {
        return VLOOKUP(A75, Premium!74:1121, 12, TRUE);
    }

    public double cell_H75() {
        return ExcelFunctions.IF(B75="", 0, (VLOOKUP(D75, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F75+G75))) + Input_AdminFee);
    }

    public double cell_I75() {
        return ExcelFunctions.IF(B75="", 0, (F75+G75-H75) * Calc_MthlyRate_Cur);
    }

    public double cell_J75() {
        return ExcelFunctions.IF(B75="", 0, Math.max(0, F75+G75-H75+I75));
    }

    public double cell_K75() {
        return ExcelFunctions.IF(B75="", 0, VLOOKUP(D75, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L75() {
        return ExcelFunctions.IF(B75="", 0, Math.max(0, J75-K75));
    }

    public double cell_M75() {
        return ExcelFunctions.IF(B75="", 0, Math.max(Input_FaceAmt, J75 * VLOOKUP(D75, Table_Rates, 4, FALSE)));
    }

    public double cell_O75() {
        return ACV;
    }

    public double cell_P75() {
        return Premium!L75;
    }

    public double cell_Q75() {
        return ExcelFunctions.IF(B75="", 0, ((VLOOKUP(D75, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O75+P75))) + Input_AdminFee);
    }

    public double cell_R75() {
        return ExcelFunctions.IF(B75="", 0, (O75+P75-Q75) * Calc_MthlyRate_Guar);
    }

    public double cell_S75() {
        return ExcelFunctions.IF(B75="", 0, Math.max(0, O75+P75-Q75+R75));
    }

    public double cell_A76() {
        return EDATE(A75, 1);
    }

    public double cell_B76() {
        return ExcelFunctions.IF(B75>=Input_ProjYears, "", ExcelFunctions.IF(C75=12, B75+1, B75));
    }

    public double cell_C76() {
        return ExcelFunctions.IF(C75=12, 1, C75+1);
    }

    public double cell_D76() {
        return Input_IssueAge+(B76-1);
    }

    public double cell_F76() {
        return ExcelFunctions.IF(B76="", 0, J75);
    }

    public double cell_G76() {
        return VLOOKUP(A76, Premium!75:1122, 12, TRUE);
    }

    public double cell_H76() {
        return ExcelFunctions.IF(B76="", 0, (VLOOKUP(D76, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F76+G76))) + Input_AdminFee);
    }

    public double cell_I76() {
        return ExcelFunctions.IF(B76="", 0, (F76+G76-H76) * Calc_MthlyRate_Cur);
    }

    public double cell_J76() {
        return ExcelFunctions.IF(B76="", 0, Math.max(0, F76+G76-H76+I76));
    }

    public double cell_K76() {
        return ExcelFunctions.IF(B76="", 0, VLOOKUP(D76, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L76() {
        return ExcelFunctions.IF(B76="", 0, Math.max(0, J76-K76));
    }

    public double cell_M76() {
        return ExcelFunctions.IF(B76="", 0, Math.max(Input_FaceAmt, J76 * VLOOKUP(D76, Table_Rates, 4, FALSE)));
    }

    public double cell_O76() {
        return ACV;
    }

    public double cell_P76() {
        return Premium!L76;
    }

    public double cell_Q76() {
        return ExcelFunctions.IF(B76="", 0, ((VLOOKUP(D76, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O76+P76))) + Input_AdminFee);
    }

    public double cell_R76() {
        return ExcelFunctions.IF(B76="", 0, (O76+P76-Q76) * Calc_MthlyRate_Guar);
    }

    public double cell_S76() {
        return ExcelFunctions.IF(B76="", 0, Math.max(0, O76+P76-Q76+R76));
    }

    public double cell_A77() {
        return EDATE(A76, 1);
    }

    public double cell_B77() {
        return ExcelFunctions.IF(B76>=Input_ProjYears, "", ExcelFunctions.IF(C76=12, B76+1, B76));
    }

    public double cell_C77() {
        return ExcelFunctions.IF(C76=12, 1, C76+1);
    }

    public double cell_D77() {
        return Input_IssueAge+(B77-1);
    }

    public double cell_F77() {
        return ExcelFunctions.IF(B77="", 0, J76);
    }

    public double cell_G77() {
        return VLOOKUP(A77, Premium!76:1123, 12, TRUE);
    }

    public double cell_H77() {
        return ExcelFunctions.IF(B77="", 0, (VLOOKUP(D77, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F77+G77))) + Input_AdminFee);
    }

    public double cell_I77() {
        return ExcelFunctions.IF(B77="", 0, (F77+G77-H77) * Calc_MthlyRate_Cur);
    }

    public double cell_J77() {
        return ExcelFunctions.IF(B77="", 0, Math.max(0, F77+G77-H77+I77));
    }

    public double cell_K77() {
        return ExcelFunctions.IF(B77="", 0, VLOOKUP(D77, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L77() {
        return ExcelFunctions.IF(B77="", 0, Math.max(0, J77-K77));
    }

    public double cell_M77() {
        return ExcelFunctions.IF(B77="", 0, Math.max(Input_FaceAmt, J77 * VLOOKUP(D77, Table_Rates, 4, FALSE)));
    }

    public double cell_O77() {
        return ACV;
    }

    public double cell_P77() {
        return Premium!L77;
    }

    public double cell_Q77() {
        return ExcelFunctions.IF(B77="", 0, ((VLOOKUP(D77, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O77+P77))) + Input_AdminFee);
    }

    public double cell_R77() {
        return ExcelFunctions.IF(B77="", 0, (O77+P77-Q77) * Calc_MthlyRate_Guar);
    }

    public double cell_S77() {
        return ExcelFunctions.IF(B77="", 0, Math.max(0, O77+P77-Q77+R77));
    }

    public double cell_A78() {
        return EDATE(A77, 1);
    }

    public double cell_B78() {
        return ExcelFunctions.IF(B77>=Input_ProjYears, "", ExcelFunctions.IF(C77=12, B77+1, B77));
    }

    public double cell_C78() {
        return ExcelFunctions.IF(C77=12, 1, C77+1);
    }

    public double cell_D78() {
        return Input_IssueAge+(B78-1);
    }

    public double cell_F78() {
        return ExcelFunctions.IF(B78="", 0, J77);
    }

    public double cell_G78() {
        return VLOOKUP(A78, Premium!77:1124, 12, TRUE);
    }

    public double cell_H78() {
        return ExcelFunctions.IF(B78="", 0, (VLOOKUP(D78, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F78+G78))) + Input_AdminFee);
    }

    public double cell_I78() {
        return ExcelFunctions.IF(B78="", 0, (F78+G78-H78) * Calc_MthlyRate_Cur);
    }

    public double cell_J78() {
        return ExcelFunctions.IF(B78="", 0, Math.max(0, F78+G78-H78+I78));
    }

    public double cell_K78() {
        return ExcelFunctions.IF(B78="", 0, VLOOKUP(D78, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L78() {
        return ExcelFunctions.IF(B78="", 0, Math.max(0, J78-K78));
    }

    public double cell_M78() {
        return ExcelFunctions.IF(B78="", 0, Math.max(Input_FaceAmt, J78 * VLOOKUP(D78, Table_Rates, 4, FALSE)));
    }

    public double cell_O78() {
        return ACV;
    }

    public double cell_P78() {
        return Premium!L78;
    }

    public double cell_Q78() {
        return ExcelFunctions.IF(B78="", 0, ((VLOOKUP(D78, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O78+P78))) + Input_AdminFee);
    }

    public double cell_R78() {
        return ExcelFunctions.IF(B78="", 0, (O78+P78-Q78) * Calc_MthlyRate_Guar);
    }

    public double cell_S78() {
        return ExcelFunctions.IF(B78="", 0, Math.max(0, O78+P78-Q78+R78));
    }

    public double cell_A79() {
        return EDATE(A78, 1);
    }

    public double cell_B79() {
        return ExcelFunctions.IF(B78>=Input_ProjYears, "", ExcelFunctions.IF(C78=12, B78+1, B78));
    }

    public double cell_C79() {
        return ExcelFunctions.IF(C78=12, 1, C78+1);
    }

    public double cell_D79() {
        return Input_IssueAge+(B79-1);
    }

    public double cell_F79() {
        return ExcelFunctions.IF(B79="", 0, J78);
    }

    public double cell_G79() {
        return VLOOKUP(A79, Premium!78:1125, 12, TRUE);
    }

    public double cell_H79() {
        return ExcelFunctions.IF(B79="", 0, (VLOOKUP(D79, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F79+G79))) + Input_AdminFee);
    }

    public double cell_I79() {
        return ExcelFunctions.IF(B79="", 0, (F79+G79-H79) * Calc_MthlyRate_Cur);
    }

    public double cell_J79() {
        return ExcelFunctions.IF(B79="", 0, Math.max(0, F79+G79-H79+I79));
    }

    public double cell_K79() {
        return ExcelFunctions.IF(B79="", 0, VLOOKUP(D79, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L79() {
        return ExcelFunctions.IF(B79="", 0, Math.max(0, J79-K79));
    }

    public double cell_M79() {
        return ExcelFunctions.IF(B79="", 0, Math.max(Input_FaceAmt, J79 * VLOOKUP(D79, Table_Rates, 4, FALSE)));
    }

    public double cell_O79() {
        return ACV;
    }

    public double cell_P79() {
        return Premium!L79;
    }

    public double cell_Q79() {
        return ExcelFunctions.IF(B79="", 0, ((VLOOKUP(D79, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O79+P79))) + Input_AdminFee);
    }

    public double cell_R79() {
        return ExcelFunctions.IF(B79="", 0, (O79+P79-Q79) * Calc_MthlyRate_Guar);
    }

    public double cell_S79() {
        return ExcelFunctions.IF(B79="", 0, Math.max(0, O79+P79-Q79+R79));
    }

    public double cell_A80() {
        return EDATE(A79, 1);
    }

    public double cell_B80() {
        return ExcelFunctions.IF(B79>=Input_ProjYears, "", ExcelFunctions.IF(C79=12, B79+1, B79));
    }

    public double cell_C80() {
        return ExcelFunctions.IF(C79=12, 1, C79+1);
    }

    public double cell_D80() {
        return Input_IssueAge+(B80-1);
    }

    public double cell_F80() {
        return ExcelFunctions.IF(B80="", 0, J79);
    }

    public double cell_G80() {
        return VLOOKUP(A80, Premium!79:1126, 12, TRUE);
    }

    public double cell_H80() {
        return ExcelFunctions.IF(B80="", 0, (VLOOKUP(D80, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F80+G80))) + Input_AdminFee);
    }

    public double cell_I80() {
        return ExcelFunctions.IF(B80="", 0, (F80+G80-H80) * Calc_MthlyRate_Cur);
    }

    public double cell_J80() {
        return ExcelFunctions.IF(B80="", 0, Math.max(0, F80+G80-H80+I80));
    }

    public double cell_K80() {
        return ExcelFunctions.IF(B80="", 0, VLOOKUP(D80, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L80() {
        return ExcelFunctions.IF(B80="", 0, Math.max(0, J80-K80));
    }

    public double cell_M80() {
        return ExcelFunctions.IF(B80="", 0, Math.max(Input_FaceAmt, J80 * VLOOKUP(D80, Table_Rates, 4, FALSE)));
    }

    public double cell_O80() {
        return ACV;
    }

    public double cell_P80() {
        return Premium!L80;
    }

    public double cell_Q80() {
        return ExcelFunctions.IF(B80="", 0, ((VLOOKUP(D80, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O80+P80))) + Input_AdminFee);
    }

    public double cell_R80() {
        return ExcelFunctions.IF(B80="", 0, (O80+P80-Q80) * Calc_MthlyRate_Guar);
    }

    public double cell_S80() {
        return ExcelFunctions.IF(B80="", 0, Math.max(0, O80+P80-Q80+R80));
    }

    public double cell_A81() {
        return EDATE(A80, 1);
    }

    public double cell_B81() {
        return ExcelFunctions.IF(B80>=Input_ProjYears, "", ExcelFunctions.IF(C80=12, B80+1, B80));
    }

    public double cell_C81() {
        return ExcelFunctions.IF(C80=12, 1, C80+1);
    }

    public double cell_D81() {
        return Input_IssueAge+(B81-1);
    }

    public double cell_F81() {
        return ExcelFunctions.IF(B81="", 0, J80);
    }

    public double cell_G81() {
        return VLOOKUP(A81, Premium!80:1127, 12, TRUE);
    }

    public double cell_H81() {
        return ExcelFunctions.IF(B81="", 0, (VLOOKUP(D81, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F81+G81))) + Input_AdminFee);
    }

    public double cell_I81() {
        return ExcelFunctions.IF(B81="", 0, (F81+G81-H81) * Calc_MthlyRate_Cur);
    }

    public double cell_J81() {
        return ExcelFunctions.IF(B81="", 0, Math.max(0, F81+G81-H81+I81));
    }

    public double cell_K81() {
        return ExcelFunctions.IF(B81="", 0, VLOOKUP(D81, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L81() {
        return ExcelFunctions.IF(B81="", 0, Math.max(0, J81-K81));
    }

    public double cell_M81() {
        return ExcelFunctions.IF(B81="", 0, Math.max(Input_FaceAmt, J81 * VLOOKUP(D81, Table_Rates, 4, FALSE)));
    }

    public double cell_O81() {
        return ACV;
    }

    public double cell_P81() {
        return Premium!L81;
    }

    public double cell_Q81() {
        return ExcelFunctions.IF(B81="", 0, ((VLOOKUP(D81, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O81+P81))) + Input_AdminFee);
    }

    public double cell_R81() {
        return ExcelFunctions.IF(B81="", 0, (O81+P81-Q81) * Calc_MthlyRate_Guar);
    }

    public double cell_S81() {
        return ExcelFunctions.IF(B81="", 0, Math.max(0, O81+P81-Q81+R81));
    }

    public double cell_A82() {
        return EDATE(A81, 1);
    }

    public double cell_B82() {
        return ExcelFunctions.IF(B81>=Input_ProjYears, "", ExcelFunctions.IF(C81=12, B81+1, B81));
    }

    public double cell_C82() {
        return ExcelFunctions.IF(C81=12, 1, C81+1);
    }

    public double cell_D82() {
        return Input_IssueAge+(B82-1);
    }

    public double cell_F82() {
        return ExcelFunctions.IF(B82="", 0, J81);
    }

    public double cell_G82() {
        return VLOOKUP(A82, Premium!81:1128, 12, TRUE);
    }

    public double cell_H82() {
        return ExcelFunctions.IF(B82="", 0, (VLOOKUP(D82, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F82+G82))) + Input_AdminFee);
    }

    public double cell_I82() {
        return ExcelFunctions.IF(B82="", 0, (F82+G82-H82) * Calc_MthlyRate_Cur);
    }

    public double cell_J82() {
        return ExcelFunctions.IF(B82="", 0, Math.max(0, F82+G82-H82+I82));
    }

    public double cell_K82() {
        return ExcelFunctions.IF(B82="", 0, VLOOKUP(D82, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L82() {
        return ExcelFunctions.IF(B82="", 0, Math.max(0, J82-K82));
    }

    public double cell_M82() {
        return ExcelFunctions.IF(B82="", 0, Math.max(Input_FaceAmt, J82 * VLOOKUP(D82, Table_Rates, 4, FALSE)));
    }

    public double cell_O82() {
        return ACV;
    }

    public double cell_P82() {
        return Premium!L82;
    }

    public double cell_Q82() {
        return ExcelFunctions.IF(B82="", 0, ((VLOOKUP(D82, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O82+P82))) + Input_AdminFee);
    }

    public double cell_R82() {
        return ExcelFunctions.IF(B82="", 0, (O82+P82-Q82) * Calc_MthlyRate_Guar);
    }

    public double cell_S82() {
        return ExcelFunctions.IF(B82="", 0, Math.max(0, O82+P82-Q82+R82));
    }

    public double cell_A83() {
        return EDATE(A82, 1);
    }

    public double cell_B83() {
        return ExcelFunctions.IF(B82>=Input_ProjYears, "", ExcelFunctions.IF(C82=12, B82+1, B82));
    }

    public double cell_C83() {
        return ExcelFunctions.IF(C82=12, 1, C82+1);
    }

    public double cell_D83() {
        return Input_IssueAge+(B83-1);
    }

    public double cell_F83() {
        return ExcelFunctions.IF(B83="", 0, J82);
    }

    public double cell_G83() {
        return VLOOKUP(A83, Premium!82:1129, 12, TRUE);
    }

    public double cell_H83() {
        return ExcelFunctions.IF(B83="", 0, (VLOOKUP(D83, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F83+G83))) + Input_AdminFee);
    }

    public double cell_I83() {
        return ExcelFunctions.IF(B83="", 0, (F83+G83-H83) * Calc_MthlyRate_Cur);
    }

    public double cell_J83() {
        return ExcelFunctions.IF(B83="", 0, Math.max(0, F83+G83-H83+I83));
    }

    public double cell_K83() {
        return ExcelFunctions.IF(B83="", 0, VLOOKUP(D83, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L83() {
        return ExcelFunctions.IF(B83="", 0, Math.max(0, J83-K83));
    }

    public double cell_M83() {
        return ExcelFunctions.IF(B83="", 0, Math.max(Input_FaceAmt, J83 * VLOOKUP(D83, Table_Rates, 4, FALSE)));
    }

    public double cell_O83() {
        return ACV;
    }

    public double cell_P83() {
        return Premium!L83;
    }

    public double cell_Q83() {
        return ExcelFunctions.IF(B83="", 0, ((VLOOKUP(D83, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O83+P83))) + Input_AdminFee);
    }

    public double cell_R83() {
        return ExcelFunctions.IF(B83="", 0, (O83+P83-Q83) * Calc_MthlyRate_Guar);
    }

    public double cell_S83() {
        return ExcelFunctions.IF(B83="", 0, Math.max(0, O83+P83-Q83+R83));
    }

    public double cell_A84() {
        return EDATE(A83, 1);
    }

    public double cell_B84() {
        return ExcelFunctions.IF(B83>=Input_ProjYears, "", ExcelFunctions.IF(C83=12, B83+1, B83));
    }

    public double cell_C84() {
        return ExcelFunctions.IF(C83=12, 1, C83+1);
    }

    public double cell_D84() {
        return Input_IssueAge+(B84-1);
    }

    public double cell_F84() {
        return ExcelFunctions.IF(B84="", 0, J83);
    }

    public double cell_G84() {
        return VLOOKUP(A84, Premium!83:1130, 12, TRUE);
    }

    public double cell_H84() {
        return ExcelFunctions.IF(B84="", 0, (VLOOKUP(D84, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F84+G84))) + Input_AdminFee);
    }

    public double cell_I84() {
        return ExcelFunctions.IF(B84="", 0, (F84+G84-H84) * Calc_MthlyRate_Cur);
    }

    public double cell_J84() {
        return ExcelFunctions.IF(B84="", 0, Math.max(0, F84+G84-H84+I84));
    }

    public double cell_K84() {
        return ExcelFunctions.IF(B84="", 0, VLOOKUP(D84, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L84() {
        return ExcelFunctions.IF(B84="", 0, Math.max(0, J84-K84));
    }

    public double cell_M84() {
        return ExcelFunctions.IF(B84="", 0, Math.max(Input_FaceAmt, J84 * VLOOKUP(D84, Table_Rates, 4, FALSE)));
    }

    public double cell_O84() {
        return ACV;
    }

    public double cell_P84() {
        return Premium!L84;
    }

    public double cell_Q84() {
        return ExcelFunctions.IF(B84="", 0, ((VLOOKUP(D84, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O84+P84))) + Input_AdminFee);
    }

    public double cell_R84() {
        return ExcelFunctions.IF(B84="", 0, (O84+P84-Q84) * Calc_MthlyRate_Guar);
    }

    public double cell_S84() {
        return ExcelFunctions.IF(B84="", 0, Math.max(0, O84+P84-Q84+R84));
    }

    public double cell_A85() {
        return EDATE(A84, 1);
    }

    public double cell_B85() {
        return ExcelFunctions.IF(B84>=Input_ProjYears, "", ExcelFunctions.IF(C84=12, B84+1, B84));
    }

    public double cell_C85() {
        return ExcelFunctions.IF(C84=12, 1, C84+1);
    }

    public double cell_D85() {
        return Input_IssueAge+(B85-1);
    }

    public double cell_F85() {
        return ExcelFunctions.IF(B85="", 0, J84);
    }

    public double cell_G85() {
        return VLOOKUP(A85, Premium!84:1131, 12, TRUE);
    }

    public double cell_H85() {
        return ExcelFunctions.IF(B85="", 0, (VLOOKUP(D85, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F85+G85))) + Input_AdminFee);
    }

    public double cell_I85() {
        return ExcelFunctions.IF(B85="", 0, (F85+G85-H85) * Calc_MthlyRate_Cur);
    }

    public double cell_J85() {
        return ExcelFunctions.IF(B85="", 0, Math.max(0, F85+G85-H85+I85));
    }

    public double cell_K85() {
        return ExcelFunctions.IF(B85="", 0, VLOOKUP(D85, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L85() {
        return ExcelFunctions.IF(B85="", 0, Math.max(0, J85-K85));
    }

    public double cell_M85() {
        return ExcelFunctions.IF(B85="", 0, Math.max(Input_FaceAmt, J85 * VLOOKUP(D85, Table_Rates, 4, FALSE)));
    }

    public double cell_O85() {
        return ACV;
    }

    public double cell_P85() {
        return Premium!L85;
    }

    public double cell_Q85() {
        return ExcelFunctions.IF(B85="", 0, ((VLOOKUP(D85, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O85+P85))) + Input_AdminFee);
    }

    public double cell_R85() {
        return ExcelFunctions.IF(B85="", 0, (O85+P85-Q85) * Calc_MthlyRate_Guar);
    }

    public double cell_S85() {
        return ExcelFunctions.IF(B85="", 0, Math.max(0, O85+P85-Q85+R85));
    }

    public double cell_A86() {
        return EDATE(A85, 1);
    }

    public double cell_B86() {
        return ExcelFunctions.IF(B85>=Input_ProjYears, "", ExcelFunctions.IF(C85=12, B85+1, B85));
    }

    public double cell_C86() {
        return ExcelFunctions.IF(C85=12, 1, C85+1);
    }

    public double cell_D86() {
        return Input_IssueAge+(B86-1);
    }

    public double cell_F86() {
        return ExcelFunctions.IF(B86="", 0, J85);
    }

    public double cell_G86() {
        return VLOOKUP(A86, Premium!85:1132, 12, TRUE);
    }

    public double cell_H86() {
        return ExcelFunctions.IF(B86="", 0, (VLOOKUP(D86, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F86+G86))) + Input_AdminFee);
    }

    public double cell_I86() {
        return ExcelFunctions.IF(B86="", 0, (F86+G86-H86) * Calc_MthlyRate_Cur);
    }

    public double cell_J86() {
        return ExcelFunctions.IF(B86="", 0, Math.max(0, F86+G86-H86+I86));
    }

    public double cell_K86() {
        return ExcelFunctions.IF(B86="", 0, VLOOKUP(D86, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L86() {
        return ExcelFunctions.IF(B86="", 0, Math.max(0, J86-K86));
    }

    public double cell_M86() {
        return ExcelFunctions.IF(B86="", 0, Math.max(Input_FaceAmt, J86 * VLOOKUP(D86, Table_Rates, 4, FALSE)));
    }

    public double cell_O86() {
        return ACV;
    }

    public double cell_P86() {
        return Premium!L86;
    }

    public double cell_Q86() {
        return ExcelFunctions.IF(B86="", 0, ((VLOOKUP(D86, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O86+P86))) + Input_AdminFee);
    }

    public double cell_R86() {
        return ExcelFunctions.IF(B86="", 0, (O86+P86-Q86) * Calc_MthlyRate_Guar);
    }

    public double cell_S86() {
        return ExcelFunctions.IF(B86="", 0, Math.max(0, O86+P86-Q86+R86));
    }

    public double cell_A87() {
        return EDATE(A86, 1);
    }

    public double cell_B87() {
        return ExcelFunctions.IF(B86>=Input_ProjYears, "", ExcelFunctions.IF(C86=12, B86+1, B86));
    }

    public double cell_C87() {
        return ExcelFunctions.IF(C86=12, 1, C86+1);
    }

    public double cell_D87() {
        return Input_IssueAge+(B87-1);
    }

    public double cell_F87() {
        return ExcelFunctions.IF(B87="", 0, J86);
    }

    public double cell_G87() {
        return VLOOKUP(A87, Premium!86:1133, 12, TRUE);
    }

    public double cell_H87() {
        return ExcelFunctions.IF(B87="", 0, (VLOOKUP(D87, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F87+G87))) + Input_AdminFee);
    }

    public double cell_I87() {
        return ExcelFunctions.IF(B87="", 0, (F87+G87-H87) * Calc_MthlyRate_Cur);
    }

    public double cell_J87() {
        return ExcelFunctions.IF(B87="", 0, Math.max(0, F87+G87-H87+I87));
    }

    public double cell_K87() {
        return ExcelFunctions.IF(B87="", 0, VLOOKUP(D87, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L87() {
        return ExcelFunctions.IF(B87="", 0, Math.max(0, J87-K87));
    }

    public double cell_M87() {
        return ExcelFunctions.IF(B87="", 0, Math.max(Input_FaceAmt, J87 * VLOOKUP(D87, Table_Rates, 4, FALSE)));
    }

    public double cell_O87() {
        return ACV;
    }

    public double cell_P87() {
        return Premium!L87;
    }

    public double cell_Q87() {
        return ExcelFunctions.IF(B87="", 0, ((VLOOKUP(D87, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O87+P87))) + Input_AdminFee);
    }

    public double cell_R87() {
        return ExcelFunctions.IF(B87="", 0, (O87+P87-Q87) * Calc_MthlyRate_Guar);
    }

    public double cell_S87() {
        return ExcelFunctions.IF(B87="", 0, Math.max(0, O87+P87-Q87+R87));
    }

    public double cell_A88() {
        return EDATE(A87, 1);
    }

    public double cell_B88() {
        return ExcelFunctions.IF(B87>=Input_ProjYears, "", ExcelFunctions.IF(C87=12, B87+1, B87));
    }

    public double cell_C88() {
        return ExcelFunctions.IF(C87=12, 1, C87+1);
    }

    public double cell_D88() {
        return Input_IssueAge+(B88-1);
    }

    public double cell_F88() {
        return ExcelFunctions.IF(B88="", 0, J87);
    }

    public double cell_G88() {
        return VLOOKUP(A88, Premium!87:1134, 12, TRUE);
    }

    public double cell_H88() {
        return ExcelFunctions.IF(B88="", 0, (VLOOKUP(D88, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F88+G88))) + Input_AdminFee);
    }

    public double cell_I88() {
        return ExcelFunctions.IF(B88="", 0, (F88+G88-H88) * Calc_MthlyRate_Cur);
    }

    public double cell_J88() {
        return ExcelFunctions.IF(B88="", 0, Math.max(0, F88+G88-H88+I88));
    }

    public double cell_K88() {
        return ExcelFunctions.IF(B88="", 0, VLOOKUP(D88, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L88() {
        return ExcelFunctions.IF(B88="", 0, Math.max(0, J88-K88));
    }

    public double cell_M88() {
        return ExcelFunctions.IF(B88="", 0, Math.max(Input_FaceAmt, J88 * VLOOKUP(D88, Table_Rates, 4, FALSE)));
    }

    public double cell_O88() {
        return ACV;
    }

    public double cell_P88() {
        return Premium!L88;
    }

    public double cell_Q88() {
        return ExcelFunctions.IF(B88="", 0, ((VLOOKUP(D88, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O88+P88))) + Input_AdminFee);
    }

    public double cell_R88() {
        return ExcelFunctions.IF(B88="", 0, (O88+P88-Q88) * Calc_MthlyRate_Guar);
    }

    public double cell_S88() {
        return ExcelFunctions.IF(B88="", 0, Math.max(0, O88+P88-Q88+R88));
    }

    public double cell_A89() {
        return EDATE(A88, 1);
    }

    public double cell_B89() {
        return ExcelFunctions.IF(B88>=Input_ProjYears, "", ExcelFunctions.IF(C88=12, B88+1, B88));
    }

    public double cell_C89() {
        return ExcelFunctions.IF(C88=12, 1, C88+1);
    }

    public double cell_D89() {
        return Input_IssueAge+(B89-1);
    }

    public double cell_F89() {
        return ExcelFunctions.IF(B89="", 0, J88);
    }

    public double cell_G89() {
        return VLOOKUP(A89, Premium!88:1135, 12, TRUE);
    }

    public double cell_H89() {
        return ExcelFunctions.IF(B89="", 0, (VLOOKUP(D89, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F89+G89))) + Input_AdminFee);
    }

    public double cell_I89() {
        return ExcelFunctions.IF(B89="", 0, (F89+G89-H89) * Calc_MthlyRate_Cur);
    }

    public double cell_J89() {
        return ExcelFunctions.IF(B89="", 0, Math.max(0, F89+G89-H89+I89));
    }

    public double cell_K89() {
        return ExcelFunctions.IF(B89="", 0, VLOOKUP(D89, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L89() {
        return ExcelFunctions.IF(B89="", 0, Math.max(0, J89-K89));
    }

    public double cell_M89() {
        return ExcelFunctions.IF(B89="", 0, Math.max(Input_FaceAmt, J89 * VLOOKUP(D89, Table_Rates, 4, FALSE)));
    }

    public double cell_O89() {
        return ACV;
    }

    public double cell_P89() {
        return Premium!L89;
    }

    public double cell_Q89() {
        return ExcelFunctions.IF(B89="", 0, ((VLOOKUP(D89, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O89+P89))) + Input_AdminFee);
    }

    public double cell_R89() {
        return ExcelFunctions.IF(B89="", 0, (O89+P89-Q89) * Calc_MthlyRate_Guar);
    }

    public double cell_S89() {
        return ExcelFunctions.IF(B89="", 0, Math.max(0, O89+P89-Q89+R89));
    }

    public double cell_A90() {
        return EDATE(A89, 1);
    }

    public double cell_B90() {
        return ExcelFunctions.IF(B89>=Input_ProjYears, "", ExcelFunctions.IF(C89=12, B89+1, B89));
    }

    public double cell_C90() {
        return ExcelFunctions.IF(C89=12, 1, C89+1);
    }

    public double cell_D90() {
        return Input_IssueAge+(B90-1);
    }

    public double cell_F90() {
        return ExcelFunctions.IF(B90="", 0, J89);
    }

    public double cell_G90() {
        return VLOOKUP(A90, Premium!89:1136, 12, TRUE);
    }

    public double cell_H90() {
        return ExcelFunctions.IF(B90="", 0, (VLOOKUP(D90, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F90+G90))) + Input_AdminFee);
    }

    public double cell_I90() {
        return ExcelFunctions.IF(B90="", 0, (F90+G90-H90) * Calc_MthlyRate_Cur);
    }

    public double cell_J90() {
        return ExcelFunctions.IF(B90="", 0, Math.max(0, F90+G90-H90+I90));
    }

    public double cell_K90() {
        return ExcelFunctions.IF(B90="", 0, VLOOKUP(D90, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L90() {
        return ExcelFunctions.IF(B90="", 0, Math.max(0, J90-K90));
    }

    public double cell_M90() {
        return ExcelFunctions.IF(B90="", 0, Math.max(Input_FaceAmt, J90 * VLOOKUP(D90, Table_Rates, 4, FALSE)));
    }

    public double cell_O90() {
        return ACV;
    }

    public double cell_P90() {
        return Premium!L90;
    }

    public double cell_Q90() {
        return ExcelFunctions.IF(B90="", 0, ((VLOOKUP(D90, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O90+P90))) + Input_AdminFee);
    }

    public double cell_R90() {
        return ExcelFunctions.IF(B90="", 0, (O90+P90-Q90) * Calc_MthlyRate_Guar);
    }

    public double cell_S90() {
        return ExcelFunctions.IF(B90="", 0, Math.max(0, O90+P90-Q90+R90));
    }

    public double cell_A91() {
        return EDATE(A90, 1);
    }

    public double cell_B91() {
        return ExcelFunctions.IF(B90>=Input_ProjYears, "", ExcelFunctions.IF(C90=12, B90+1, B90));
    }

    public double cell_C91() {
        return ExcelFunctions.IF(C90=12, 1, C90+1);
    }

    public double cell_D91() {
        return Input_IssueAge+(B91-1);
    }

    public double cell_F91() {
        return ExcelFunctions.IF(B91="", 0, J90);
    }

    public double cell_G91() {
        return VLOOKUP(A91, Premium!90:1137, 12, TRUE);
    }

    public double cell_H91() {
        return ExcelFunctions.IF(B91="", 0, (VLOOKUP(D91, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F91+G91))) + Input_AdminFee);
    }

    public double cell_I91() {
        return ExcelFunctions.IF(B91="", 0, (F91+G91-H91) * Calc_MthlyRate_Cur);
    }

    public double cell_J91() {
        return ExcelFunctions.IF(B91="", 0, Math.max(0, F91+G91-H91+I91));
    }

    public double cell_K91() {
        return ExcelFunctions.IF(B91="", 0, VLOOKUP(D91, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L91() {
        return ExcelFunctions.IF(B91="", 0, Math.max(0, J91-K91));
    }

    public double cell_M91() {
        return ExcelFunctions.IF(B91="", 0, Math.max(Input_FaceAmt, J91 * VLOOKUP(D91, Table_Rates, 4, FALSE)));
    }

    public double cell_O91() {
        return ACV;
    }

    public double cell_P91() {
        return Premium!L91;
    }

    public double cell_Q91() {
        return ExcelFunctions.IF(B91="", 0, ((VLOOKUP(D91, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O91+P91))) + Input_AdminFee);
    }

    public double cell_R91() {
        return ExcelFunctions.IF(B91="", 0, (O91+P91-Q91) * Calc_MthlyRate_Guar);
    }

    public double cell_S91() {
        return ExcelFunctions.IF(B91="", 0, Math.max(0, O91+P91-Q91+R91));
    }

    public double cell_A92() {
        return EDATE(A91, 1);
    }

    public double cell_B92() {
        return ExcelFunctions.IF(B91>=Input_ProjYears, "", ExcelFunctions.IF(C91=12, B91+1, B91));
    }

    public double cell_C92() {
        return ExcelFunctions.IF(C91=12, 1, C91+1);
    }

    public double cell_D92() {
        return Input_IssueAge+(B92-1);
    }

    public double cell_F92() {
        return ExcelFunctions.IF(B92="", 0, J91);
    }

    public double cell_G92() {
        return VLOOKUP(A92, Premium!91:1138, 12, TRUE);
    }

    public double cell_H92() {
        return ExcelFunctions.IF(B92="", 0, (VLOOKUP(D92, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F92+G92))) + Input_AdminFee);
    }

    public double cell_I92() {
        return ExcelFunctions.IF(B92="", 0, (F92+G92-H92) * Calc_MthlyRate_Cur);
    }

    public double cell_J92() {
        return ExcelFunctions.IF(B92="", 0, Math.max(0, F92+G92-H92+I92));
    }

    public double cell_K92() {
        return ExcelFunctions.IF(B92="", 0, VLOOKUP(D92, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L92() {
        return ExcelFunctions.IF(B92="", 0, Math.max(0, J92-K92));
    }

    public double cell_M92() {
        return ExcelFunctions.IF(B92="", 0, Math.max(Input_FaceAmt, J92 * VLOOKUP(D92, Table_Rates, 4, FALSE)));
    }

    public double cell_O92() {
        return ACV;
    }

    public double cell_P92() {
        return Premium!L92;
    }

    public double cell_Q92() {
        return ExcelFunctions.IF(B92="", 0, ((VLOOKUP(D92, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O92+P92))) + Input_AdminFee);
    }

    public double cell_R92() {
        return ExcelFunctions.IF(B92="", 0, (O92+P92-Q92) * Calc_MthlyRate_Guar);
    }

    public double cell_S92() {
        return ExcelFunctions.IF(B92="", 0, Math.max(0, O92+P92-Q92+R92));
    }

    public double cell_A93() {
        return EDATE(A92, 1);
    }

    public double cell_B93() {
        return ExcelFunctions.IF(B92>=Input_ProjYears, "", ExcelFunctions.IF(C92=12, B92+1, B92));
    }

    public double cell_C93() {
        return ExcelFunctions.IF(C92=12, 1, C92+1);
    }

    public double cell_D93() {
        return Input_IssueAge+(B93-1);
    }

    public double cell_F93() {
        return ExcelFunctions.IF(B93="", 0, J92);
    }

    public double cell_G93() {
        return VLOOKUP(A93, Premium!92:1139, 12, TRUE);
    }

    public double cell_H93() {
        return ExcelFunctions.IF(B93="", 0, (VLOOKUP(D93, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F93+G93))) + Input_AdminFee);
    }

    public double cell_I93() {
        return ExcelFunctions.IF(B93="", 0, (F93+G93-H93) * Calc_MthlyRate_Cur);
    }

    public double cell_J93() {
        return ExcelFunctions.IF(B93="", 0, Math.max(0, F93+G93-H93+I93));
    }

    public double cell_K93() {
        return ExcelFunctions.IF(B93="", 0, VLOOKUP(D93, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L93() {
        return ExcelFunctions.IF(B93="", 0, Math.max(0, J93-K93));
    }

    public double cell_M93() {
        return ExcelFunctions.IF(B93="", 0, Math.max(Input_FaceAmt, J93 * VLOOKUP(D93, Table_Rates, 4, FALSE)));
    }

    public double cell_O93() {
        return ACV;
    }

    public double cell_P93() {
        return Premium!L93;
    }

    public double cell_Q93() {
        return ExcelFunctions.IF(B93="", 0, ((VLOOKUP(D93, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O93+P93))) + Input_AdminFee);
    }

    public double cell_R93() {
        return ExcelFunctions.IF(B93="", 0, (O93+P93-Q93) * Calc_MthlyRate_Guar);
    }

    public double cell_S93() {
        return ExcelFunctions.IF(B93="", 0, Math.max(0, O93+P93-Q93+R93));
    }

    public double cell_A94() {
        return EDATE(A93, 1);
    }

    public double cell_B94() {
        return ExcelFunctions.IF(B93>=Input_ProjYears, "", ExcelFunctions.IF(C93=12, B93+1, B93));
    }

    public double cell_C94() {
        return ExcelFunctions.IF(C93=12, 1, C93+1);
    }

    public double cell_D94() {
        return Input_IssueAge+(B94-1);
    }

    public double cell_F94() {
        return ExcelFunctions.IF(B94="", 0, J93);
    }

    public double cell_G94() {
        return VLOOKUP(A94, Premium!93:1140, 12, TRUE);
    }

    public double cell_H94() {
        return ExcelFunctions.IF(B94="", 0, (VLOOKUP(D94, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F94+G94))) + Input_AdminFee);
    }

    public double cell_I94() {
        return ExcelFunctions.IF(B94="", 0, (F94+G94-H94) * Calc_MthlyRate_Cur);
    }

    public double cell_J94() {
        return ExcelFunctions.IF(B94="", 0, Math.max(0, F94+G94-H94+I94));
    }

    public double cell_K94() {
        return ExcelFunctions.IF(B94="", 0, VLOOKUP(D94, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L94() {
        return ExcelFunctions.IF(B94="", 0, Math.max(0, J94-K94));
    }

    public double cell_M94() {
        return ExcelFunctions.IF(B94="", 0, Math.max(Input_FaceAmt, J94 * VLOOKUP(D94, Table_Rates, 4, FALSE)));
    }

    public double cell_O94() {
        return ACV;
    }

    public double cell_P94() {
        return Premium!L94;
    }

    public double cell_Q94() {
        return ExcelFunctions.IF(B94="", 0, ((VLOOKUP(D94, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O94+P94))) + Input_AdminFee);
    }

    public double cell_R94() {
        return ExcelFunctions.IF(B94="", 0, (O94+P94-Q94) * Calc_MthlyRate_Guar);
    }

    public double cell_S94() {
        return ExcelFunctions.IF(B94="", 0, Math.max(0, O94+P94-Q94+R94));
    }

    public double cell_A95() {
        return EDATE(A94, 1);
    }

    public double cell_B95() {
        return ExcelFunctions.IF(B94>=Input_ProjYears, "", ExcelFunctions.IF(C94=12, B94+1, B94));
    }

    public double cell_C95() {
        return ExcelFunctions.IF(C94=12, 1, C94+1);
    }

    public double cell_D95() {
        return Input_IssueAge+(B95-1);
    }

    public double cell_F95() {
        return ExcelFunctions.IF(B95="", 0, J94);
    }

    public double cell_G95() {
        return VLOOKUP(A95, Premium!94:1141, 12, TRUE);
    }

    public double cell_H95() {
        return ExcelFunctions.IF(B95="", 0, (VLOOKUP(D95, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F95+G95))) + Input_AdminFee);
    }

    public double cell_I95() {
        return ExcelFunctions.IF(B95="", 0, (F95+G95-H95) * Calc_MthlyRate_Cur);
    }

    public double cell_J95() {
        return ExcelFunctions.IF(B95="", 0, Math.max(0, F95+G95-H95+I95));
    }

    public double cell_K95() {
        return ExcelFunctions.IF(B95="", 0, VLOOKUP(D95, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L95() {
        return ExcelFunctions.IF(B95="", 0, Math.max(0, J95-K95));
    }

    public double cell_M95() {
        return ExcelFunctions.IF(B95="", 0, Math.max(Input_FaceAmt, J95 * VLOOKUP(D95, Table_Rates, 4, FALSE)));
    }

    public double cell_O95() {
        return ACV;
    }

    public double cell_P95() {
        return Premium!L95;
    }

    public double cell_Q95() {
        return ExcelFunctions.IF(B95="", 0, ((VLOOKUP(D95, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O95+P95))) + Input_AdminFee);
    }

    public double cell_R95() {
        return ExcelFunctions.IF(B95="", 0, (O95+P95-Q95) * Calc_MthlyRate_Guar);
    }

    public double cell_S95() {
        return ExcelFunctions.IF(B95="", 0, Math.max(0, O95+P95-Q95+R95));
    }

    public double cell_A96() {
        return EDATE(A95, 1);
    }

    public double cell_B96() {
        return ExcelFunctions.IF(B95>=Input_ProjYears, "", ExcelFunctions.IF(C95=12, B95+1, B95));
    }

    public double cell_C96() {
        return ExcelFunctions.IF(C95=12, 1, C95+1);
    }

    public double cell_D96() {
        return Input_IssueAge+(B96-1);
    }

    public double cell_F96() {
        return ExcelFunctions.IF(B96="", 0, J95);
    }

    public double cell_G96() {
        return VLOOKUP(A96, Premium!95:1142, 12, TRUE);
    }

    public double cell_H96() {
        return ExcelFunctions.IF(B96="", 0, (VLOOKUP(D96, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F96+G96))) + Input_AdminFee);
    }

    public double cell_I96() {
        return ExcelFunctions.IF(B96="", 0, (F96+G96-H96) * Calc_MthlyRate_Cur);
    }

    public double cell_J96() {
        return ExcelFunctions.IF(B96="", 0, Math.max(0, F96+G96-H96+I96));
    }

    public double cell_K96() {
        return ExcelFunctions.IF(B96="", 0, VLOOKUP(D96, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L96() {
        return ExcelFunctions.IF(B96="", 0, Math.max(0, J96-K96));
    }

    public double cell_M96() {
        return ExcelFunctions.IF(B96="", 0, Math.max(Input_FaceAmt, J96 * VLOOKUP(D96, Table_Rates, 4, FALSE)));
    }

    public double cell_O96() {
        return ACV;
    }

    public double cell_P96() {
        return Premium!L96;
    }

    public double cell_Q96() {
        return ExcelFunctions.IF(B96="", 0, ((VLOOKUP(D96, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O96+P96))) + Input_AdminFee);
    }

    public double cell_R96() {
        return ExcelFunctions.IF(B96="", 0, (O96+P96-Q96) * Calc_MthlyRate_Guar);
    }

    public double cell_S96() {
        return ExcelFunctions.IF(B96="", 0, Math.max(0, O96+P96-Q96+R96));
    }

    public double cell_A97() {
        return EDATE(A96, 1);
    }

    public double cell_B97() {
        return ExcelFunctions.IF(B96>=Input_ProjYears, "", ExcelFunctions.IF(C96=12, B96+1, B96));
    }

    public double cell_C97() {
        return ExcelFunctions.IF(C96=12, 1, C96+1);
    }

    public double cell_D97() {
        return Input_IssueAge+(B97-1);
    }

    public double cell_F97() {
        return ExcelFunctions.IF(B97="", 0, J96);
    }

    public double cell_G97() {
        return VLOOKUP(A97, Premium!96:1143, 12, TRUE);
    }

    public double cell_H97() {
        return ExcelFunctions.IF(B97="", 0, (VLOOKUP(D97, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F97+G97))) + Input_AdminFee);
    }

    public double cell_I97() {
        return ExcelFunctions.IF(B97="", 0, (F97+G97-H97) * Calc_MthlyRate_Cur);
    }

    public double cell_J97() {
        return ExcelFunctions.IF(B97="", 0, Math.max(0, F97+G97-H97+I97));
    }

    public double cell_K97() {
        return ExcelFunctions.IF(B97="", 0, VLOOKUP(D97, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L97() {
        return ExcelFunctions.IF(B97="", 0, Math.max(0, J97-K97));
    }

    public double cell_M97() {
        return ExcelFunctions.IF(B97="", 0, Math.max(Input_FaceAmt, J97 * VLOOKUP(D97, Table_Rates, 4, FALSE)));
    }

    public double cell_O97() {
        return ACV;
    }

    public double cell_P97() {
        return Premium!L97;
    }

    public double cell_Q97() {
        return ExcelFunctions.IF(B97="", 0, ((VLOOKUP(D97, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O97+P97))) + Input_AdminFee);
    }

    public double cell_R97() {
        return ExcelFunctions.IF(B97="", 0, (O97+P97-Q97) * Calc_MthlyRate_Guar);
    }

    public double cell_S97() {
        return ExcelFunctions.IF(B97="", 0, Math.max(0, O97+P97-Q97+R97));
    }

    public double cell_A98() {
        return EDATE(A97, 1);
    }

    public double cell_B98() {
        return ExcelFunctions.IF(B97>=Input_ProjYears, "", ExcelFunctions.IF(C97=12, B97+1, B97));
    }

    public double cell_C98() {
        return ExcelFunctions.IF(C97=12, 1, C97+1);
    }

    public double cell_D98() {
        return Input_IssueAge+(B98-1);
    }

    public double cell_F98() {
        return ExcelFunctions.IF(B98="", 0, J97);
    }

    public double cell_G98() {
        return VLOOKUP(A98, Premium!97:1144, 12, TRUE);
    }

    public double cell_H98() {
        return ExcelFunctions.IF(B98="", 0, (VLOOKUP(D98, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F98+G98))) + Input_AdminFee);
    }

    public double cell_I98() {
        return ExcelFunctions.IF(B98="", 0, (F98+G98-H98) * Calc_MthlyRate_Cur);
    }

    public double cell_J98() {
        return ExcelFunctions.IF(B98="", 0, Math.max(0, F98+G98-H98+I98));
    }

    public double cell_K98() {
        return ExcelFunctions.IF(B98="", 0, VLOOKUP(D98, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L98() {
        return ExcelFunctions.IF(B98="", 0, Math.max(0, J98-K98));
    }

    public double cell_M98() {
        return ExcelFunctions.IF(B98="", 0, Math.max(Input_FaceAmt, J98 * VLOOKUP(D98, Table_Rates, 4, FALSE)));
    }

    public double cell_O98() {
        return ACV;
    }

    public double cell_P98() {
        return Premium!L98;
    }

    public double cell_Q98() {
        return ExcelFunctions.IF(B98="", 0, ((VLOOKUP(D98, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O98+P98))) + Input_AdminFee);
    }

    public double cell_R98() {
        return ExcelFunctions.IF(B98="", 0, (O98+P98-Q98) * Calc_MthlyRate_Guar);
    }

    public double cell_S98() {
        return ExcelFunctions.IF(B98="", 0, Math.max(0, O98+P98-Q98+R98));
    }

    public double cell_A99() {
        return EDATE(A98, 1);
    }

    public double cell_B99() {
        return ExcelFunctions.IF(B98>=Input_ProjYears, "", ExcelFunctions.IF(C98=12, B98+1, B98));
    }

    public double cell_C99() {
        return ExcelFunctions.IF(C98=12, 1, C98+1);
    }

    public double cell_D99() {
        return Input_IssueAge+(B99-1);
    }

    public double cell_F99() {
        return ExcelFunctions.IF(B99="", 0, J98);
    }

    public double cell_G99() {
        return VLOOKUP(A99, Premium!98:1145, 12, TRUE);
    }

    public double cell_H99() {
        return ExcelFunctions.IF(B99="", 0, (VLOOKUP(D99, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F99+G99))) + Input_AdminFee);
    }

    public double cell_I99() {
        return ExcelFunctions.IF(B99="", 0, (F99+G99-H99) * Calc_MthlyRate_Cur);
    }

    public double cell_J99() {
        return ExcelFunctions.IF(B99="", 0, Math.max(0, F99+G99-H99+I99));
    }

    public double cell_K99() {
        return ExcelFunctions.IF(B99="", 0, VLOOKUP(D99, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L99() {
        return ExcelFunctions.IF(B99="", 0, Math.max(0, J99-K99));
    }

    public double cell_M99() {
        return ExcelFunctions.IF(B99="", 0, Math.max(Input_FaceAmt, J99 * VLOOKUP(D99, Table_Rates, 4, FALSE)));
    }

    public double cell_O99() {
        return ACV;
    }

    public double cell_P99() {
        return Premium!L99;
    }

    public double cell_Q99() {
        return ExcelFunctions.IF(B99="", 0, ((VLOOKUP(D99, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O99+P99))) + Input_AdminFee);
    }

    public double cell_R99() {
        return ExcelFunctions.IF(B99="", 0, (O99+P99-Q99) * Calc_MthlyRate_Guar);
    }

    public double cell_S99() {
        return ExcelFunctions.IF(B99="", 0, Math.max(0, O99+P99-Q99+R99));
    }

    public double cell_A100() {
        return EDATE(A99, 1);
    }

    public double cell_B100() {
        return ExcelFunctions.IF(B99>=Input_ProjYears, "", ExcelFunctions.IF(C99=12, B99+1, B99));
    }

    public double cell_C100() {
        return ExcelFunctions.IF(C99=12, 1, C99+1);
    }

    public double cell_D100() {
        return Input_IssueAge+(B100-1);
    }

    public double cell_F100() {
        return ExcelFunctions.IF(B100="", 0, J99);
    }

    public double cell_G100() {
        return VLOOKUP(A100, Premium!99:1146, 12, TRUE);
    }

    public double cell_H100() {
        return ExcelFunctions.IF(B100="", 0, (VLOOKUP(D100, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F100+G100))) + Input_AdminFee);
    }

    public double cell_I100() {
        return ExcelFunctions.IF(B100="", 0, (F100+G100-H100) * Calc_MthlyRate_Cur);
    }

    public double cell_J100() {
        return ExcelFunctions.IF(B100="", 0, Math.max(0, F100+G100-H100+I100));
    }

    public double cell_K100() {
        return ExcelFunctions.IF(B100="", 0, VLOOKUP(D100, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L100() {
        return ExcelFunctions.IF(B100="", 0, Math.max(0, J100-K100));
    }

    public double cell_M100() {
        return ExcelFunctions.IF(B100="", 0, Math.max(Input_FaceAmt, J100 * VLOOKUP(D100, Table_Rates, 4, FALSE)));
    }

    public double cell_O100() {
        return ACV;
    }

    public double cell_P100() {
        return Premium!L100;
    }

    public double cell_Q100() {
        return ExcelFunctions.IF(B100="", 0, ((VLOOKUP(D100, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O100+P100))) + Input_AdminFee);
    }

    public double cell_R100() {
        return ExcelFunctions.IF(B100="", 0, (O100+P100-Q100) * Calc_MthlyRate_Guar);
    }

    public double cell_S100() {
        return ExcelFunctions.IF(B100="", 0, Math.max(0, O100+P100-Q100+R100));
    }

    public double cell_A101() {
        return EDATE(A100, 1);
    }

    public double cell_B101() {
        return ExcelFunctions.IF(B100>=Input_ProjYears, "", ExcelFunctions.IF(C100=12, B100+1, B100));
    }

    public double cell_C101() {
        return ExcelFunctions.IF(C100=12, 1, C100+1);
    }

    public double cell_D101() {
        return Input_IssueAge+(B101-1);
    }

    public double cell_F101() {
        return ExcelFunctions.IF(B101="", 0, J100);
    }

    public double cell_G101() {
        return VLOOKUP(A101, Premium!100:1147, 12, TRUE);
    }

    public double cell_H101() {
        return ExcelFunctions.IF(B101="", 0, (VLOOKUP(D101, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F101+G101))) + Input_AdminFee);
    }

    public double cell_I101() {
        return ExcelFunctions.IF(B101="", 0, (F101+G101-H101) * Calc_MthlyRate_Cur);
    }

    public double cell_J101() {
        return ExcelFunctions.IF(B101="", 0, Math.max(0, F101+G101-H101+I101));
    }

    public double cell_K101() {
        return ExcelFunctions.IF(B101="", 0, VLOOKUP(D101, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L101() {
        return ExcelFunctions.IF(B101="", 0, Math.max(0, J101-K101));
    }

    public double cell_M101() {
        return ExcelFunctions.IF(B101="", 0, Math.max(Input_FaceAmt, J101 * VLOOKUP(D101, Table_Rates, 4, FALSE)));
    }

    public double cell_O101() {
        return ACV;
    }

    public double cell_P101() {
        return Premium!L101;
    }

    public double cell_Q101() {
        return ExcelFunctions.IF(B101="", 0, ((VLOOKUP(D101, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O101+P101))) + Input_AdminFee);
    }

    public double cell_R101() {
        return ExcelFunctions.IF(B101="", 0, (O101+P101-Q101) * Calc_MthlyRate_Guar);
    }

    public double cell_S101() {
        return ExcelFunctions.IF(B101="", 0, Math.max(0, O101+P101-Q101+R101));
    }

    public double cell_A102() {
        return EDATE(A101, 1);
    }

    public double cell_B102() {
        return ExcelFunctions.IF(B101>=Input_ProjYears, "", ExcelFunctions.IF(C101=12, B101+1, B101));
    }

    public double cell_C102() {
        return ExcelFunctions.IF(C101=12, 1, C101+1);
    }

    public double cell_D102() {
        return Input_IssueAge+(B102-1);
    }

    public double cell_F102() {
        return ExcelFunctions.IF(B102="", 0, J101);
    }

    public double cell_G102() {
        return VLOOKUP(A102, Premium!101:1148, 12, TRUE);
    }

    public double cell_H102() {
        return ExcelFunctions.IF(B102="", 0, (VLOOKUP(D102, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F102+G102))) + Input_AdminFee);
    }

    public double cell_I102() {
        return ExcelFunctions.IF(B102="", 0, (F102+G102-H102) * Calc_MthlyRate_Cur);
    }

    public double cell_J102() {
        return ExcelFunctions.IF(B102="", 0, Math.max(0, F102+G102-H102+I102));
    }

    public double cell_K102() {
        return ExcelFunctions.IF(B102="", 0, VLOOKUP(D102, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L102() {
        return ExcelFunctions.IF(B102="", 0, Math.max(0, J102-K102));
    }

    public double cell_M102() {
        return ExcelFunctions.IF(B102="", 0, Math.max(Input_FaceAmt, J102 * VLOOKUP(D102, Table_Rates, 4, FALSE)));
    }

    public double cell_O102() {
        return ACV;
    }

    public double cell_P102() {
        return Premium!L102;
    }

    public double cell_Q102() {
        return ExcelFunctions.IF(B102="", 0, ((VLOOKUP(D102, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O102+P102))) + Input_AdminFee);
    }

    public double cell_R102() {
        return ExcelFunctions.IF(B102="", 0, (O102+P102-Q102) * Calc_MthlyRate_Guar);
    }

    public double cell_S102() {
        return ExcelFunctions.IF(B102="", 0, Math.max(0, O102+P102-Q102+R102));
    }

    public double cell_A103() {
        return EDATE(A102, 1);
    }

    public double cell_B103() {
        return ExcelFunctions.IF(B102>=Input_ProjYears, "", ExcelFunctions.IF(C102=12, B102+1, B102));
    }

    public double cell_C103() {
        return ExcelFunctions.IF(C102=12, 1, C102+1);
    }

    public double cell_D103() {
        return Input_IssueAge+(B103-1);
    }

    public double cell_F103() {
        return ExcelFunctions.IF(B103="", 0, J102);
    }

    public double cell_G103() {
        return VLOOKUP(A103, Premium!102:1149, 12, TRUE);
    }

    public double cell_H103() {
        return ExcelFunctions.IF(B103="", 0, (VLOOKUP(D103, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F103+G103))) + Input_AdminFee);
    }

    public double cell_I103() {
        return ExcelFunctions.IF(B103="", 0, (F103+G103-H103) * Calc_MthlyRate_Cur);
    }

    public double cell_J103() {
        return ExcelFunctions.IF(B103="", 0, Math.max(0, F103+G103-H103+I103));
    }

    public double cell_K103() {
        return ExcelFunctions.IF(B103="", 0, VLOOKUP(D103, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L103() {
        return ExcelFunctions.IF(B103="", 0, Math.max(0, J103-K103));
    }

    public double cell_M103() {
        return ExcelFunctions.IF(B103="", 0, Math.max(Input_FaceAmt, J103 * VLOOKUP(D103, Table_Rates, 4, FALSE)));
    }

    public double cell_O103() {
        return ACV;
    }

    public double cell_P103() {
        return Premium!L103;
    }

    public double cell_Q103() {
        return ExcelFunctions.IF(B103="", 0, ((VLOOKUP(D103, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O103+P103))) + Input_AdminFee);
    }

    public double cell_R103() {
        return ExcelFunctions.IF(B103="", 0, (O103+P103-Q103) * Calc_MthlyRate_Guar);
    }

    public double cell_S103() {
        return ExcelFunctions.IF(B103="", 0, Math.max(0, O103+P103-Q103+R103));
    }

    public double cell_A104() {
        return EDATE(A103, 1);
    }

    public double cell_B104() {
        return ExcelFunctions.IF(B103>=Input_ProjYears, "", ExcelFunctions.IF(C103=12, B103+1, B103));
    }

    public double cell_C104() {
        return ExcelFunctions.IF(C103=12, 1, C103+1);
    }

    public double cell_D104() {
        return Input_IssueAge+(B104-1);
    }

    public double cell_F104() {
        return ExcelFunctions.IF(B104="", 0, J103);
    }

    public double cell_G104() {
        return VLOOKUP(A104, Premium!103:1150, 12, TRUE);
    }

    public double cell_H104() {
        return ExcelFunctions.IF(B104="", 0, (VLOOKUP(D104, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F104+G104))) + Input_AdminFee);
    }

    public double cell_I104() {
        return ExcelFunctions.IF(B104="", 0, (F104+G104-H104) * Calc_MthlyRate_Cur);
    }

    public double cell_J104() {
        return ExcelFunctions.IF(B104="", 0, Math.max(0, F104+G104-H104+I104));
    }

    public double cell_K104() {
        return ExcelFunctions.IF(B104="", 0, VLOOKUP(D104, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L104() {
        return ExcelFunctions.IF(B104="", 0, Math.max(0, J104-K104));
    }

    public double cell_M104() {
        return ExcelFunctions.IF(B104="", 0, Math.max(Input_FaceAmt, J104 * VLOOKUP(D104, Table_Rates, 4, FALSE)));
    }

    public double cell_O104() {
        return ACV;
    }

    public double cell_P104() {
        return Premium!L104;
    }

    public double cell_Q104() {
        return ExcelFunctions.IF(B104="", 0, ((VLOOKUP(D104, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O104+P104))) + Input_AdminFee);
    }

    public double cell_R104() {
        return ExcelFunctions.IF(B104="", 0, (O104+P104-Q104) * Calc_MthlyRate_Guar);
    }

    public double cell_S104() {
        return ExcelFunctions.IF(B104="", 0, Math.max(0, O104+P104-Q104+R104));
    }

    public double cell_A105() {
        return EDATE(A104, 1);
    }

    public double cell_B105() {
        return ExcelFunctions.IF(B104>=Input_ProjYears, "", ExcelFunctions.IF(C104=12, B104+1, B104));
    }

    public double cell_C105() {
        return ExcelFunctions.IF(C104=12, 1, C104+1);
    }

    public double cell_D105() {
        return Input_IssueAge+(B105-1);
    }

    public double cell_F105() {
        return ExcelFunctions.IF(B105="", 0, J104);
    }

    public double cell_G105() {
        return VLOOKUP(A105, Premium!104:1151, 12, TRUE);
    }

    public double cell_H105() {
        return ExcelFunctions.IF(B105="", 0, (VLOOKUP(D105, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F105+G105))) + Input_AdminFee);
    }

    public double cell_I105() {
        return ExcelFunctions.IF(B105="", 0, (F105+G105-H105) * Calc_MthlyRate_Cur);
    }

    public double cell_J105() {
        return ExcelFunctions.IF(B105="", 0, Math.max(0, F105+G105-H105+I105));
    }

    public double cell_K105() {
        return ExcelFunctions.IF(B105="", 0, VLOOKUP(D105, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L105() {
        return ExcelFunctions.IF(B105="", 0, Math.max(0, J105-K105));
    }

    public double cell_M105() {
        return ExcelFunctions.IF(B105="", 0, Math.max(Input_FaceAmt, J105 * VLOOKUP(D105, Table_Rates, 4, FALSE)));
    }

    public double cell_O105() {
        return ACV;
    }

    public double cell_P105() {
        return Premium!L105;
    }

    public double cell_Q105() {
        return ExcelFunctions.IF(B105="", 0, ((VLOOKUP(D105, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O105+P105))) + Input_AdminFee);
    }

    public double cell_R105() {
        return ExcelFunctions.IF(B105="", 0, (O105+P105-Q105) * Calc_MthlyRate_Guar);
    }

    public double cell_S105() {
        return ExcelFunctions.IF(B105="", 0, Math.max(0, O105+P105-Q105+R105));
    }

    public double cell_A106() {
        return EDATE(A105, 1);
    }

    public double cell_B106() {
        return ExcelFunctions.IF(B105>=Input_ProjYears, "", ExcelFunctions.IF(C105=12, B105+1, B105));
    }

    public double cell_C106() {
        return ExcelFunctions.IF(C105=12, 1, C105+1);
    }

    public double cell_D106() {
        return Input_IssueAge+(B106-1);
    }

    public double cell_F106() {
        return ExcelFunctions.IF(B106="", 0, J105);
    }

    public double cell_G106() {
        return VLOOKUP(A106, Premium!105:1152, 12, TRUE);
    }

    public double cell_H106() {
        return ExcelFunctions.IF(B106="", 0, (VLOOKUP(D106, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F106+G106))) + Input_AdminFee);
    }

    public double cell_I106() {
        return ExcelFunctions.IF(B106="", 0, (F106+G106-H106) * Calc_MthlyRate_Cur);
    }

    public double cell_J106() {
        return ExcelFunctions.IF(B106="", 0, Math.max(0, F106+G106-H106+I106));
    }

    public double cell_K106() {
        return ExcelFunctions.IF(B106="", 0, VLOOKUP(D106, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L106() {
        return ExcelFunctions.IF(B106="", 0, Math.max(0, J106-K106));
    }

    public double cell_M106() {
        return ExcelFunctions.IF(B106="", 0, Math.max(Input_FaceAmt, J106 * VLOOKUP(D106, Table_Rates, 4, FALSE)));
    }

    public double cell_O106() {
        return ACV;
    }

    public double cell_P106() {
        return Premium!L106;
    }

    public double cell_Q106() {
        return ExcelFunctions.IF(B106="", 0, ((VLOOKUP(D106, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O106+P106))) + Input_AdminFee);
    }

    public double cell_R106() {
        return ExcelFunctions.IF(B106="", 0, (O106+P106-Q106) * Calc_MthlyRate_Guar);
    }

    public double cell_S106() {
        return ExcelFunctions.IF(B106="", 0, Math.max(0, O106+P106-Q106+R106));
    }

    public double cell_A107() {
        return EDATE(A106, 1);
    }

    public double cell_B107() {
        return ExcelFunctions.IF(B106>=Input_ProjYears, "", ExcelFunctions.IF(C106=12, B106+1, B106));
    }

    public double cell_C107() {
        return ExcelFunctions.IF(C106=12, 1, C106+1);
    }

    public double cell_D107() {
        return Input_IssueAge+(B107-1);
    }

    public double cell_F107() {
        return ExcelFunctions.IF(B107="", 0, J106);
    }

    public double cell_G107() {
        return VLOOKUP(A107, Premium!106:1153, 12, TRUE);
    }

    public double cell_H107() {
        return ExcelFunctions.IF(B107="", 0, (VLOOKUP(D107, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F107+G107))) + Input_AdminFee);
    }

    public double cell_I107() {
        return ExcelFunctions.IF(B107="", 0, (F107+G107-H107) * Calc_MthlyRate_Cur);
    }

    public double cell_J107() {
        return ExcelFunctions.IF(B107="", 0, Math.max(0, F107+G107-H107+I107));
    }

    public double cell_K107() {
        return ExcelFunctions.IF(B107="", 0, VLOOKUP(D107, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L107() {
        return ExcelFunctions.IF(B107="", 0, Math.max(0, J107-K107));
    }

    public double cell_M107() {
        return ExcelFunctions.IF(B107="", 0, Math.max(Input_FaceAmt, J107 * VLOOKUP(D107, Table_Rates, 4, FALSE)));
    }

    public double cell_O107() {
        return ACV;
    }

    public double cell_P107() {
        return Premium!L107;
    }

    public double cell_Q107() {
        return ExcelFunctions.IF(B107="", 0, ((VLOOKUP(D107, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O107+P107))) + Input_AdminFee);
    }

    public double cell_R107() {
        return ExcelFunctions.IF(B107="", 0, (O107+P107-Q107) * Calc_MthlyRate_Guar);
    }

    public double cell_S107() {
        return ExcelFunctions.IF(B107="", 0, Math.max(0, O107+P107-Q107+R107));
    }

    public double cell_A108() {
        return EDATE(A107, 1);
    }

    public double cell_B108() {
        return ExcelFunctions.IF(B107>=Input_ProjYears, "", ExcelFunctions.IF(C107=12, B107+1, B107));
    }

    public double cell_C108() {
        return ExcelFunctions.IF(C107=12, 1, C107+1);
    }

    public double cell_D108() {
        return Input_IssueAge+(B108-1);
    }

    public double cell_F108() {
        return ExcelFunctions.IF(B108="", 0, J107);
    }

    public double cell_G108() {
        return VLOOKUP(A108, Premium!107:1154, 12, TRUE);
    }

    public double cell_H108() {
        return ExcelFunctions.IF(B108="", 0, (VLOOKUP(D108, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F108+G108))) + Input_AdminFee);
    }

    public double cell_I108() {
        return ExcelFunctions.IF(B108="", 0, (F108+G108-H108) * Calc_MthlyRate_Cur);
    }

    public double cell_J108() {
        return ExcelFunctions.IF(B108="", 0, Math.max(0, F108+G108-H108+I108));
    }

    public double cell_K108() {
        return ExcelFunctions.IF(B108="", 0, VLOOKUP(D108, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L108() {
        return ExcelFunctions.IF(B108="", 0, Math.max(0, J108-K108));
    }

    public double cell_M108() {
        return ExcelFunctions.IF(B108="", 0, Math.max(Input_FaceAmt, J108 * VLOOKUP(D108, Table_Rates, 4, FALSE)));
    }

    public double cell_O108() {
        return ACV;
    }

    public double cell_P108() {
        return Premium!L108;
    }

    public double cell_Q108() {
        return ExcelFunctions.IF(B108="", 0, ((VLOOKUP(D108, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O108+P108))) + Input_AdminFee);
    }

    public double cell_R108() {
        return ExcelFunctions.IF(B108="", 0, (O108+P108-Q108) * Calc_MthlyRate_Guar);
    }

    public double cell_S108() {
        return ExcelFunctions.IF(B108="", 0, Math.max(0, O108+P108-Q108+R108));
    }

    public double cell_A109() {
        return EDATE(A108, 1);
    }

    public double cell_B109() {
        return ExcelFunctions.IF(B108>=Input_ProjYears, "", ExcelFunctions.IF(C108=12, B108+1, B108));
    }

    public double cell_C109() {
        return ExcelFunctions.IF(C108=12, 1, C108+1);
    }

    public double cell_D109() {
        return Input_IssueAge+(B109-1);
    }

    public double cell_F109() {
        return ExcelFunctions.IF(B109="", 0, J108);
    }

    public double cell_G109() {
        return VLOOKUP(A109, Premium!108:1155, 12, TRUE);
    }

    public double cell_H109() {
        return ExcelFunctions.IF(B109="", 0, (VLOOKUP(D109, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F109+G109))) + Input_AdminFee);
    }

    public double cell_I109() {
        return ExcelFunctions.IF(B109="", 0, (F109+G109-H109) * Calc_MthlyRate_Cur);
    }

    public double cell_J109() {
        return ExcelFunctions.IF(B109="", 0, Math.max(0, F109+G109-H109+I109));
    }

    public double cell_K109() {
        return ExcelFunctions.IF(B109="", 0, VLOOKUP(D109, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L109() {
        return ExcelFunctions.IF(B109="", 0, Math.max(0, J109-K109));
    }

    public double cell_M109() {
        return ExcelFunctions.IF(B109="", 0, Math.max(Input_FaceAmt, J109 * VLOOKUP(D109, Table_Rates, 4, FALSE)));
    }

    public double cell_O109() {
        return ACV;
    }

    public double cell_P109() {
        return Premium!L109;
    }

    public double cell_Q109() {
        return ExcelFunctions.IF(B109="", 0, ((VLOOKUP(D109, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O109+P109))) + Input_AdminFee);
    }

    public double cell_R109() {
        return ExcelFunctions.IF(B109="", 0, (O109+P109-Q109) * Calc_MthlyRate_Guar);
    }

    public double cell_S109() {
        return ExcelFunctions.IF(B109="", 0, Math.max(0, O109+P109-Q109+R109));
    }

    public double cell_A110() {
        return EDATE(A109, 1);
    }

    public double cell_B110() {
        return ExcelFunctions.IF(B109>=Input_ProjYears, "", ExcelFunctions.IF(C109=12, B109+1, B109));
    }

    public double cell_C110() {
        return ExcelFunctions.IF(C109=12, 1, C109+1);
    }

    public double cell_D110() {
        return Input_IssueAge+(B110-1);
    }

    public double cell_F110() {
        return ExcelFunctions.IF(B110="", 0, J109);
    }

    public double cell_G110() {
        return VLOOKUP(A110, Premium!109:1156, 12, TRUE);
    }

    public double cell_H110() {
        return ExcelFunctions.IF(B110="", 0, (VLOOKUP(D110, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F110+G110))) + Input_AdminFee);
    }

    public double cell_I110() {
        return ExcelFunctions.IF(B110="", 0, (F110+G110-H110) * Calc_MthlyRate_Cur);
    }

    public double cell_J110() {
        return ExcelFunctions.IF(B110="", 0, Math.max(0, F110+G110-H110+I110));
    }

    public double cell_K110() {
        return ExcelFunctions.IF(B110="", 0, VLOOKUP(D110, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L110() {
        return ExcelFunctions.IF(B110="", 0, Math.max(0, J110-K110));
    }

    public double cell_M110() {
        return ExcelFunctions.IF(B110="", 0, Math.max(Input_FaceAmt, J110 * VLOOKUP(D110, Table_Rates, 4, FALSE)));
    }

    public double cell_O110() {
        return ACV;
    }

    public double cell_P110() {
        return Premium!L110;
    }

    public double cell_Q110() {
        return ExcelFunctions.IF(B110="", 0, ((VLOOKUP(D110, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O110+P110))) + Input_AdminFee);
    }

    public double cell_R110() {
        return ExcelFunctions.IF(B110="", 0, (O110+P110-Q110) * Calc_MthlyRate_Guar);
    }

    public double cell_S110() {
        return ExcelFunctions.IF(B110="", 0, Math.max(0, O110+P110-Q110+R110));
    }

    public double cell_A111() {
        return EDATE(A110, 1);
    }

    public double cell_B111() {
        return ExcelFunctions.IF(B110>=Input_ProjYears, "", ExcelFunctions.IF(C110=12, B110+1, B110));
    }

    public double cell_C111() {
        return ExcelFunctions.IF(C110=12, 1, C110+1);
    }

    public double cell_D111() {
        return Input_IssueAge+(B111-1);
    }

    public double cell_F111() {
        return ExcelFunctions.IF(B111="", 0, J110);
    }

    public double cell_G111() {
        return VLOOKUP(A111, Premium!110:1157, 12, TRUE);
    }

    public double cell_H111() {
        return ExcelFunctions.IF(B111="", 0, (VLOOKUP(D111, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F111+G111))) + Input_AdminFee);
    }

    public double cell_I111() {
        return ExcelFunctions.IF(B111="", 0, (F111+G111-H111) * Calc_MthlyRate_Cur);
    }

    public double cell_J111() {
        return ExcelFunctions.IF(B111="", 0, Math.max(0, F111+G111-H111+I111));
    }

    public double cell_K111() {
        return ExcelFunctions.IF(B111="", 0, VLOOKUP(D111, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L111() {
        return ExcelFunctions.IF(B111="", 0, Math.max(0, J111-K111));
    }

    public double cell_M111() {
        return ExcelFunctions.IF(B111="", 0, Math.max(Input_FaceAmt, J111 * VLOOKUP(D111, Table_Rates, 4, FALSE)));
    }

    public double cell_O111() {
        return ACV;
    }

    public double cell_P111() {
        return Premium!L111;
    }

    public double cell_Q111() {
        return ExcelFunctions.IF(B111="", 0, ((VLOOKUP(D111, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O111+P111))) + Input_AdminFee);
    }

    public double cell_R111() {
        return ExcelFunctions.IF(B111="", 0, (O111+P111-Q111) * Calc_MthlyRate_Guar);
    }

    public double cell_S111() {
        return ExcelFunctions.IF(B111="", 0, Math.max(0, O111+P111-Q111+R111));
    }

    public double cell_A112() {
        return EDATE(A111, 1);
    }

    public double cell_B112() {
        return ExcelFunctions.IF(B111>=Input_ProjYears, "", ExcelFunctions.IF(C111=12, B111+1, B111));
    }

    public double cell_C112() {
        return ExcelFunctions.IF(C111=12, 1, C111+1);
    }

    public double cell_D112() {
        return Input_IssueAge+(B112-1);
    }

    public double cell_F112() {
        return ExcelFunctions.IF(B112="", 0, J111);
    }

    public double cell_G112() {
        return VLOOKUP(A112, Premium!111:1158, 12, TRUE);
    }

    public double cell_H112() {
        return ExcelFunctions.IF(B112="", 0, (VLOOKUP(D112, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F112+G112))) + Input_AdminFee);
    }

    public double cell_I112() {
        return ExcelFunctions.IF(B112="", 0, (F112+G112-H112) * Calc_MthlyRate_Cur);
    }

    public double cell_J112() {
        return ExcelFunctions.IF(B112="", 0, Math.max(0, F112+G112-H112+I112));
    }

    public double cell_K112() {
        return ExcelFunctions.IF(B112="", 0, VLOOKUP(D112, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L112() {
        return ExcelFunctions.IF(B112="", 0, Math.max(0, J112-K112));
    }

    public double cell_M112() {
        return ExcelFunctions.IF(B112="", 0, Math.max(Input_FaceAmt, J112 * VLOOKUP(D112, Table_Rates, 4, FALSE)));
    }

    public double cell_O112() {
        return ACV;
    }

    public double cell_P112() {
        return Premium!L112;
    }

    public double cell_Q112() {
        return ExcelFunctions.IF(B112="", 0, ((VLOOKUP(D112, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O112+P112))) + Input_AdminFee);
    }

    public double cell_R112() {
        return ExcelFunctions.IF(B112="", 0, (O112+P112-Q112) * Calc_MthlyRate_Guar);
    }

    public double cell_S112() {
        return ExcelFunctions.IF(B112="", 0, Math.max(0, O112+P112-Q112+R112));
    }

    public double cell_A113() {
        return EDATE(A112, 1);
    }

    public double cell_B113() {
        return ExcelFunctions.IF(B112>=Input_ProjYears, "", ExcelFunctions.IF(C112=12, B112+1, B112));
    }

    public double cell_C113() {
        return ExcelFunctions.IF(C112=12, 1, C112+1);
    }

    public double cell_D113() {
        return Input_IssueAge+(B113-1);
    }

    public double cell_F113() {
        return ExcelFunctions.IF(B113="", 0, J112);
    }

    public double cell_G113() {
        return VLOOKUP(A113, Premium!112:1159, 12, TRUE);
    }

    public double cell_H113() {
        return ExcelFunctions.IF(B113="", 0, (VLOOKUP(D113, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F113+G113))) + Input_AdminFee);
    }

    public double cell_I113() {
        return ExcelFunctions.IF(B113="", 0, (F113+G113-H113) * Calc_MthlyRate_Cur);
    }

    public double cell_J113() {
        return ExcelFunctions.IF(B113="", 0, Math.max(0, F113+G113-H113+I113));
    }

    public double cell_K113() {
        return ExcelFunctions.IF(B113="", 0, VLOOKUP(D113, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L113() {
        return ExcelFunctions.IF(B113="", 0, Math.max(0, J113-K113));
    }

    public double cell_M113() {
        return ExcelFunctions.IF(B113="", 0, Math.max(Input_FaceAmt, J113 * VLOOKUP(D113, Table_Rates, 4, FALSE)));
    }

    public double cell_O113() {
        return ACV;
    }

    public double cell_P113() {
        return Premium!L113;
    }

    public double cell_Q113() {
        return ExcelFunctions.IF(B113="", 0, ((VLOOKUP(D113, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O113+P113))) + Input_AdminFee);
    }

    public double cell_R113() {
        return ExcelFunctions.IF(B113="", 0, (O113+P113-Q113) * Calc_MthlyRate_Guar);
    }

    public double cell_S113() {
        return ExcelFunctions.IF(B113="", 0, Math.max(0, O113+P113-Q113+R113));
    }

    public double cell_A114() {
        return EDATE(A113, 1);
    }

    public double cell_B114() {
        return ExcelFunctions.IF(B113>=Input_ProjYears, "", ExcelFunctions.IF(C113=12, B113+1, B113));
    }

    public double cell_C114() {
        return ExcelFunctions.IF(C113=12, 1, C113+1);
    }

    public double cell_D114() {
        return Input_IssueAge+(B114-1);
    }

    public double cell_F114() {
        return ExcelFunctions.IF(B114="", 0, J113);
    }

    public double cell_G114() {
        return VLOOKUP(A114, Premium!113:1160, 12, TRUE);
    }

    public double cell_H114() {
        return ExcelFunctions.IF(B114="", 0, (VLOOKUP(D114, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F114+G114))) + Input_AdminFee);
    }

    public double cell_I114() {
        return ExcelFunctions.IF(B114="", 0, (F114+G114-H114) * Calc_MthlyRate_Cur);
    }

    public double cell_J114() {
        return ExcelFunctions.IF(B114="", 0, Math.max(0, F114+G114-H114+I114));
    }

    public double cell_K114() {
        return ExcelFunctions.IF(B114="", 0, VLOOKUP(D114, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L114() {
        return ExcelFunctions.IF(B114="", 0, Math.max(0, J114-K114));
    }

    public double cell_M114() {
        return ExcelFunctions.IF(B114="", 0, Math.max(Input_FaceAmt, J114 * VLOOKUP(D114, Table_Rates, 4, FALSE)));
    }

    public double cell_O114() {
        return ACV;
    }

    public double cell_P114() {
        return Premium!L114;
    }

    public double cell_Q114() {
        return ExcelFunctions.IF(B114="", 0, ((VLOOKUP(D114, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O114+P114))) + Input_AdminFee);
    }

    public double cell_R114() {
        return ExcelFunctions.IF(B114="", 0, (O114+P114-Q114) * Calc_MthlyRate_Guar);
    }

    public double cell_S114() {
        return ExcelFunctions.IF(B114="", 0, Math.max(0, O114+P114-Q114+R114));
    }

    public double cell_A115() {
        return EDATE(A114, 1);
    }

    public double cell_B115() {
        return ExcelFunctions.IF(B114>=Input_ProjYears, "", ExcelFunctions.IF(C114=12, B114+1, B114));
    }

    public double cell_C115() {
        return ExcelFunctions.IF(C114=12, 1, C114+1);
    }

    public double cell_D115() {
        return Input_IssueAge+(B115-1);
    }

    public double cell_F115() {
        return ExcelFunctions.IF(B115="", 0, J114);
    }

    public double cell_G115() {
        return VLOOKUP(A115, Premium!114:1161, 12, TRUE);
    }

    public double cell_H115() {
        return ExcelFunctions.IF(B115="", 0, (VLOOKUP(D115, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F115+G115))) + Input_AdminFee);
    }

    public double cell_I115() {
        return ExcelFunctions.IF(B115="", 0, (F115+G115-H115) * Calc_MthlyRate_Cur);
    }

    public double cell_J115() {
        return ExcelFunctions.IF(B115="", 0, Math.max(0, F115+G115-H115+I115));
    }

    public double cell_K115() {
        return ExcelFunctions.IF(B115="", 0, VLOOKUP(D115, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L115() {
        return ExcelFunctions.IF(B115="", 0, Math.max(0, J115-K115));
    }

    public double cell_M115() {
        return ExcelFunctions.IF(B115="", 0, Math.max(Input_FaceAmt, J115 * VLOOKUP(D115, Table_Rates, 4, FALSE)));
    }

    public double cell_O115() {
        return ACV;
    }

    public double cell_P115() {
        return Premium!L115;
    }

    public double cell_Q115() {
        return ExcelFunctions.IF(B115="", 0, ((VLOOKUP(D115, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O115+P115))) + Input_AdminFee);
    }

    public double cell_R115() {
        return ExcelFunctions.IF(B115="", 0, (O115+P115-Q115) * Calc_MthlyRate_Guar);
    }

    public double cell_S115() {
        return ExcelFunctions.IF(B115="", 0, Math.max(0, O115+P115-Q115+R115));
    }

    public double cell_A116() {
        return EDATE(A115, 1);
    }

    public double cell_B116() {
        return ExcelFunctions.IF(B115>=Input_ProjYears, "", ExcelFunctions.IF(C115=12, B115+1, B115));
    }

    public double cell_C116() {
        return ExcelFunctions.IF(C115=12, 1, C115+1);
    }

    public double cell_D116() {
        return Input_IssueAge+(B116-1);
    }

    public double cell_F116() {
        return ExcelFunctions.IF(B116="", 0, J115);
    }

    public double cell_G116() {
        return VLOOKUP(A116, Premium!115:1162, 12, TRUE);
    }

    public double cell_H116() {
        return ExcelFunctions.IF(B116="", 0, (VLOOKUP(D116, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F116+G116))) + Input_AdminFee);
    }

    public double cell_I116() {
        return ExcelFunctions.IF(B116="", 0, (F116+G116-H116) * Calc_MthlyRate_Cur);
    }

    public double cell_J116() {
        return ExcelFunctions.IF(B116="", 0, Math.max(0, F116+G116-H116+I116));
    }

    public double cell_K116() {
        return ExcelFunctions.IF(B116="", 0, VLOOKUP(D116, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L116() {
        return ExcelFunctions.IF(B116="", 0, Math.max(0, J116-K116));
    }

    public double cell_M116() {
        return ExcelFunctions.IF(B116="", 0, Math.max(Input_FaceAmt, J116 * VLOOKUP(D116, Table_Rates, 4, FALSE)));
    }

    public double cell_O116() {
        return ACV;
    }

    public double cell_P116() {
        return Premium!L116;
    }

    public double cell_Q116() {
        return ExcelFunctions.IF(B116="", 0, ((VLOOKUP(D116, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O116+P116))) + Input_AdminFee);
    }

    public double cell_R116() {
        return ExcelFunctions.IF(B116="", 0, (O116+P116-Q116) * Calc_MthlyRate_Guar);
    }

    public double cell_S116() {
        return ExcelFunctions.IF(B116="", 0, Math.max(0, O116+P116-Q116+R116));
    }

    public double cell_A117() {
        return EDATE(A116, 1);
    }

    public double cell_B117() {
        return ExcelFunctions.IF(B116>=Input_ProjYears, "", ExcelFunctions.IF(C116=12, B116+1, B116));
    }

    public double cell_C117() {
        return ExcelFunctions.IF(C116=12, 1, C116+1);
    }

    public double cell_D117() {
        return Input_IssueAge+(B117-1);
    }

    public double cell_F117() {
        return ExcelFunctions.IF(B117="", 0, J116);
    }

    public double cell_G117() {
        return VLOOKUP(A117, Premium!116:1163, 12, TRUE);
    }

    public double cell_H117() {
        return ExcelFunctions.IF(B117="", 0, (VLOOKUP(D117, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F117+G117))) + Input_AdminFee);
    }

    public double cell_I117() {
        return ExcelFunctions.IF(B117="", 0, (F117+G117-H117) * Calc_MthlyRate_Cur);
    }

    public double cell_J117() {
        return ExcelFunctions.IF(B117="", 0, Math.max(0, F117+G117-H117+I117));
    }

    public double cell_K117() {
        return ExcelFunctions.IF(B117="", 0, VLOOKUP(D117, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L117() {
        return ExcelFunctions.IF(B117="", 0, Math.max(0, J117-K117));
    }

    public double cell_M117() {
        return ExcelFunctions.IF(B117="", 0, Math.max(Input_FaceAmt, J117 * VLOOKUP(D117, Table_Rates, 4, FALSE)));
    }

    public double cell_O117() {
        return ACV;
    }

    public double cell_P117() {
        return Premium!L117;
    }

    public double cell_Q117() {
        return ExcelFunctions.IF(B117="", 0, ((VLOOKUP(D117, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O117+P117))) + Input_AdminFee);
    }

    public double cell_R117() {
        return ExcelFunctions.IF(B117="", 0, (O117+P117-Q117) * Calc_MthlyRate_Guar);
    }

    public double cell_S117() {
        return ExcelFunctions.IF(B117="", 0, Math.max(0, O117+P117-Q117+R117));
    }

    public double cell_A118() {
        return EDATE(A117, 1);
    }

    public double cell_B118() {
        return ExcelFunctions.IF(B117>=Input_ProjYears, "", ExcelFunctions.IF(C117=12, B117+1, B117));
    }

    public double cell_C118() {
        return ExcelFunctions.IF(C117=12, 1, C117+1);
    }

    public double cell_D118() {
        return Input_IssueAge+(B118-1);
    }

    public double cell_F118() {
        return ExcelFunctions.IF(B118="", 0, J117);
    }

    public double cell_G118() {
        return VLOOKUP(A118, Premium!117:1164, 12, TRUE);
    }

    public double cell_H118() {
        return ExcelFunctions.IF(B118="", 0, (VLOOKUP(D118, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F118+G118))) + Input_AdminFee);
    }

    public double cell_I118() {
        return ExcelFunctions.IF(B118="", 0, (F118+G118-H118) * Calc_MthlyRate_Cur);
    }

    public double cell_J118() {
        return ExcelFunctions.IF(B118="", 0, Math.max(0, F118+G118-H118+I118));
    }

    public double cell_K118() {
        return ExcelFunctions.IF(B118="", 0, VLOOKUP(D118, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L118() {
        return ExcelFunctions.IF(B118="", 0, Math.max(0, J118-K118));
    }

    public double cell_M118() {
        return ExcelFunctions.IF(B118="", 0, Math.max(Input_FaceAmt, J118 * VLOOKUP(D118, Table_Rates, 4, FALSE)));
    }

    public double cell_O118() {
        return ACV;
    }

    public double cell_P118() {
        return Premium!L118;
    }

    public double cell_Q118() {
        return ExcelFunctions.IF(B118="", 0, ((VLOOKUP(D118, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O118+P118))) + Input_AdminFee);
    }

    public double cell_R118() {
        return ExcelFunctions.IF(B118="", 0, (O118+P118-Q118) * Calc_MthlyRate_Guar);
    }

    public double cell_S118() {
        return ExcelFunctions.IF(B118="", 0, Math.max(0, O118+P118-Q118+R118));
    }

    public double cell_A119() {
        return EDATE(A118, 1);
    }

    public double cell_B119() {
        return ExcelFunctions.IF(B118>=Input_ProjYears, "", ExcelFunctions.IF(C118=12, B118+1, B118));
    }

    public double cell_C119() {
        return ExcelFunctions.IF(C118=12, 1, C118+1);
    }

    public double cell_D119() {
        return Input_IssueAge+(B119-1);
    }

    public double cell_F119() {
        return ExcelFunctions.IF(B119="", 0, J118);
    }

    public double cell_G119() {
        return VLOOKUP(A119, Premium!118:1165, 12, TRUE);
    }

    public double cell_H119() {
        return ExcelFunctions.IF(B119="", 0, (VLOOKUP(D119, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F119+G119))) + Input_AdminFee);
    }

    public double cell_I119() {
        return ExcelFunctions.IF(B119="", 0, (F119+G119-H119) * Calc_MthlyRate_Cur);
    }

    public double cell_J119() {
        return ExcelFunctions.IF(B119="", 0, Math.max(0, F119+G119-H119+I119));
    }

    public double cell_K119() {
        return ExcelFunctions.IF(B119="", 0, VLOOKUP(D119, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L119() {
        return ExcelFunctions.IF(B119="", 0, Math.max(0, J119-K119));
    }

    public double cell_M119() {
        return ExcelFunctions.IF(B119="", 0, Math.max(Input_FaceAmt, J119 * VLOOKUP(D119, Table_Rates, 4, FALSE)));
    }

    public double cell_O119() {
        return ACV;
    }

    public double cell_P119() {
        return Premium!L119;
    }

    public double cell_Q119() {
        return ExcelFunctions.IF(B119="", 0, ((VLOOKUP(D119, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O119+P119))) + Input_AdminFee);
    }

    public double cell_R119() {
        return ExcelFunctions.IF(B119="", 0, (O119+P119-Q119) * Calc_MthlyRate_Guar);
    }

    public double cell_S119() {
        return ExcelFunctions.IF(B119="", 0, Math.max(0, O119+P119-Q119+R119));
    }

    public double cell_A120() {
        return EDATE(A119, 1);
    }

    public double cell_B120() {
        return ExcelFunctions.IF(B119>=Input_ProjYears, "", ExcelFunctions.IF(C119=12, B119+1, B119));
    }

    public double cell_C120() {
        return ExcelFunctions.IF(C119=12, 1, C119+1);
    }

    public double cell_D120() {
        return Input_IssueAge+(B120-1);
    }

    public double cell_F120() {
        return ExcelFunctions.IF(B120="", 0, J119);
    }

    public double cell_G120() {
        return VLOOKUP(A120, Premium!119:1166, 12, TRUE);
    }

    public double cell_H120() {
        return ExcelFunctions.IF(B120="", 0, (VLOOKUP(D120, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F120+G120))) + Input_AdminFee);
    }

    public double cell_I120() {
        return ExcelFunctions.IF(B120="", 0, (F120+G120-H120) * Calc_MthlyRate_Cur);
    }

    public double cell_J120() {
        return ExcelFunctions.IF(B120="", 0, Math.max(0, F120+G120-H120+I120));
    }

    public double cell_K120() {
        return ExcelFunctions.IF(B120="", 0, VLOOKUP(D120, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L120() {
        return ExcelFunctions.IF(B120="", 0, Math.max(0, J120-K120));
    }

    public double cell_M120() {
        return ExcelFunctions.IF(B120="", 0, Math.max(Input_FaceAmt, J120 * VLOOKUP(D120, Table_Rates, 4, FALSE)));
    }

    public double cell_O120() {
        return ACV;
    }

    public double cell_P120() {
        return Premium!L120;
    }

    public double cell_Q120() {
        return ExcelFunctions.IF(B120="", 0, ((VLOOKUP(D120, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O120+P120))) + Input_AdminFee);
    }

    public double cell_R120() {
        return ExcelFunctions.IF(B120="", 0, (O120+P120-Q120) * Calc_MthlyRate_Guar);
    }

    public double cell_S120() {
        return ExcelFunctions.IF(B120="", 0, Math.max(0, O120+P120-Q120+R120));
    }

    public double cell_A121() {
        return EDATE(A120, 1);
    }

    public double cell_B121() {
        return ExcelFunctions.IF(B120>=Input_ProjYears, "", ExcelFunctions.IF(C120=12, B120+1, B120));
    }

    public double cell_C121() {
        return ExcelFunctions.IF(C120=12, 1, C120+1);
    }

    public double cell_D121() {
        return Input_IssueAge+(B121-1);
    }

    public double cell_F121() {
        return ExcelFunctions.IF(B121="", 0, J120);
    }

    public double cell_G121() {
        return VLOOKUP(A121, Premium!120:1167, 12, TRUE);
    }

    public double cell_H121() {
        return ExcelFunctions.IF(B121="", 0, (VLOOKUP(D121, Table_Rates, 2, FALSE) * Math.max(0, Input_FaceAmt - (F121+G121))) + Input_AdminFee);
    }

    public double cell_I121() {
        return ExcelFunctions.IF(B121="", 0, (F121+G121-H121) * Calc_MthlyRate_Cur);
    }

    public double cell_J121() {
        return ExcelFunctions.IF(B121="", 0, Math.max(0, F121+G121-H121+I121));
    }

    public double cell_K121() {
        return ExcelFunctions.IF(B121="", 0, VLOOKUP(D121, Table_Rates, 3, FALSE) * Input_FaceAmt);
    }

    public double cell_L121() {
        return ExcelFunctions.IF(B121="", 0, Math.max(0, J121-K121));
    }

    public double cell_M121() {
        return ExcelFunctions.IF(B121="", 0, Math.max(Input_FaceAmt, J121 * VLOOKUP(D121, Table_Rates, 4, FALSE)));
    }

    public double cell_O121() {
        return ACV;
    }

    public double cell_P121() {
        return Premium!L121;
    }

    public double cell_Q121() {
        return ExcelFunctions.IF(B121="", 0, ((VLOOKUP(D121, Table_Rates, 2, FALSE)*1.5) * Math.max(0, Input_FaceAmt - (O121+P121))) + Input_AdminFee);
    }

    public double cell_R121() {
        return ExcelFunctions.IF(B121="", 0, (O121+P121-Q121) * Calc_MthlyRate_Guar);
    }

    public double cell_S121() {
        return ExcelFunctions.IF(B121="", 0, Math.max(0, O121+P121-Q121+R121));
    }

}