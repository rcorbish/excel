package com.rc;


import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

class ExcelDataFormatter extends org.apache.poi.ss.usermodel.DataFormatter {
    public String formatCellValue(Cell cell) {
        return "Boo Ya0";
    }
    public String 	formatCellValue(Cell cell, FormulaEvaluator evaluator) {
        return "Boo Ya1";
    }
    public String 	formatCellValue(Cell cell, FormulaEvaluator evaluator, ConditionalFormattingEvaluator cfEvaluator){
        return "Boo Ya2";
    }
    public String 	formatRawCellContents(double value, int formatIndex, java.lang.String formatString){
        return "Boo Ya3";
    }
}