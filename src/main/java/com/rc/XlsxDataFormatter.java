package com.rc;


import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Date;

class XlsxDataFormatter extends DataFormatter {
    final char columnTypes[] ;
    final NumberFormat numberFormat = new DecimalFormat("0.#############");

    XlsxDataFormatter( char columnTypes[] ) {
        this.columnTypes = new char[columnTypes.length];
        System.arraycopy(this.columnTypes,0,columnTypes,0,columnTypes.length);
    }

    protected int getColumnIndex() {
        return 0;
    }

    public String formatCellValue(Cell cell) {
        return super.formatCellValue(cell);
    }

    public String formatCellValue(Cell cell, FormulaEvaluator evaluator) {
        return super.formatCellValue(cell, evaluator, null);
    }

    public String formatCellValue(Cell cell, FormulaEvaluator evaluator, ConditionalFormattingEvaluator cfEvaluator) {
        return super.formatCellValue(cell, evaluator, cfEvaluator);
    }
    public String formatRawCellContents(double value, int formatIndex, String formatString) {
        if(DateUtil.isADateFormat(formatIndex,formatString)){
            final var tf = DateTimeFormatter.ISO_INSTANT;
            Date d = DateUtil.getJavaDate(value);
            return tf.format(d.toInstant());
        }
        return numberFormat.format(value);
    }
}