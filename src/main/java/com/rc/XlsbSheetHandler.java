package com.rc;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

class XlsbSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private final StringBuilder sb = new StringBuilder();

    public void startSheet(String sheetName) {
        System.out.println( sheetName ) ;
    }

    public void endSheet() {
        System.out.println( "--------------------------------------------------------" ) ;
    }

    @Override
    public void startRow(int rowNum) {
        System.out.print( rowNum );
    }

    @Override
    public void endRow(int rowNum) {
        System.out.println();
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        System.out.print( "\t" + formattedValue );
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        if (isHeader) {
            System.out.print( "\n" + text );
        } else {
            System.out.print(  text + "\n" );
        }
    }

    @Override
    public String toString() {
        return sb.toString();
    }
}