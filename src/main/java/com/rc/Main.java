package com.rc;

import joptsimple.OptionParser;
import joptsimple.OptionSet;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

import org.xml.sax.InputSource;

import javax.xml.parsers.ParserConfigurationException;
import java.io.InputStream;


public class Main {


    static final Logger logger = LogManager.getLogger(Main.class.getName());

    public static void main(String[] args) {
        logger.log(Level.INFO, "Starting");
        Options options = new Options(args);

        try {
			openExcelBinary();
//            openExcel();
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(2);
        }
    }

    public static void openExcelBinary() throws Throwable {
        OPCPackage pkg = OPCPackage.open("excel.xlsb");
        XSSFBReader r = new XSSFBReader(pkg);
        XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(pkg);
        XSSFBStylesTable xssfbStylesTable = r.getXSSFBStylesTable();
        XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) r.getSheetsData();

        while (it.hasNext()) {
            InputStream is = it.next();
            String name = it.getSheetName();
            XlsbSheetHandler xlsbSheetHandler = new XlsbSheetHandler();
            xlsbSheetHandler.startSheet(name);
            XSSFBSheetHandler sheetHandler = new XSSFBSheetHandler(
                    is,
                    xssfbStylesTable,
                    it.getXSSFBSheetComments(),
                    sst,
                    xlsbSheetHandler,
                    new ExcelDataFormatter(),
                    false);
            sheetHandler.parse();
            xlsbSheetHandler.endSheet();
        }
    }


    public static void openExcel() throws Throwable {
        final var pkg = OPCPackage.open("excel.xlsx");
        final var strings = new ReadOnlySharedStringsTable(pkg);
        final var xssfReader = new XSSFReader(pkg);
        final var styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        final var outputTypes = new ColumnType[]{ ColumnType.DECIMAL,ColumnType.STRING,ColumnType.DATETIME,ColumnType.DECIMAL};

        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                InputSource sheetSource = new InputSource(stream);
                try {
                    final var sheetParser = XMLHelper.newXMLReader();
                    final var handler = new XlsxSheetHandler(styles, strings,outputTypes);
                    sheetParser.setContentHandler(handler);
                    sheetParser.parse(sheetSource);
                } catch (ParserConfigurationException e) {
                    throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
                }
            }
        }
    }
}


class Options {

    int port = 8111;
    String platform = null;

    public Options(String[] args) {
        OptionParser parser = new OptionParser();
        parser.accepts("port", "Port number for website - defaults to 8111")
                .withRequiredArg().ofType(Integer.class);

        parser.accepts("platform", "Preferred BLAs platform, cuda or openblas")
                .withRequiredArg().ofType(String.class);

        OptionSet os = parser.parse(args);
        if (os.has("port")) {
            port = (Integer) os.valueOf("port");
        }
    }

}

enum ColumnType {
    DECIMAL,
    DATETIME,
    BOOLEAN,
    STRING
}