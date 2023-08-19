/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.rc;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * This class handles the streaming processing of a sheet#.xml
 *  sheet part of a XSSF .xlsx file, and generates
 *  row and cell events for it.
 *
 * This allows to build functionality which reads huge files
 * without needing large amounts of main memory.
 *
 */
public class XlsxSheetHandler extends DefaultHandler {

    /**
     * These are the different kinds of cells we support.
     * We keep track of the current one between
     *  the start and end.
     */
    enum xssfDataType {
        BOOLEAN,
        ERROR,
        FORMULA,
        INLINE_STRING,
        SST_STRING,
        NUMBER,
    }

    /**
     * Table with the styles used for formatting
     */
    private final Styles stylesTable;

    /**
     * Read only access to the shared strings table, for looking
     *  up (most) string cell's contents
     */
    private final SharedStrings sharedStringsTable;


    // Set when V start element is seen
    private boolean vIsOpen;
    // Set when F start element is seen
    private boolean fIsOpen;
    // Set when an Inline String "is" is seen
    private boolean isIsOpen;

    // Set when cell start element is seen;
    // used when cell close element is seen.
    private xssfDataType nextDataType;
    private boolean isFormula ;

    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private int rowNum;
    private int nextRowNum;      // some sheets do not have rowNums, Excel can read them so we should try to handle them correctly as well
    private int previousColumnIndex;
    private String cellRef;

    // Gathers characters as they are seen.
    private final StringBuilder value = new StringBuilder(64);
    private final StringBuilder formula = new StringBuilder(64);

    private boolean hasProcessedHeaders ;

    private final ColumnType outputTypes[] ;

    private NumberFormat numberFormat = new DecimalFormat("0.#############");
    /**
     * Accepts objects needed while parsing.
     *
     * @param styles  Table of styles
     * @param strings Table of shared strings
     */
    public XlsxSheetHandler(
            Styles styles,
            SharedStrings strings,
            ColumnType outputTypes[]) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.nextDataType = xssfDataType.NUMBER;
        this.outputTypes = outputTypes;
        this.hasProcessedHeaders=false;
    }


    private boolean isTextTag(String name) {
        if("v".equals(name)) {
            // Easy, normal v text tag
            return true;
        }
        if("inlineStr".equals(name)) {
            // Easy inline string
            return true;
        }
        if("t".equals(name) && isIsOpen) {
            // Inline string <is><t>...</t></is> pair
            return true;
        }
        // It isn't a text tag
        return false;
    }

    @Override
    @SuppressWarnings("unused")
    public void startElement(String uri, String localName, String qName,
                             Attributes attributes) throws SAXException {

        if (uri != null && ! uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if (isTextTag(localName)) {
            vIsOpen = true;
            // Clear contents cache
            if (!isIsOpen) {
                value.setLength(0);
            }
        } else if ("is".equals(localName)) {
            // Inline string outer tag
            isIsOpen = true;
        } else if ("f".equals(localName)) {
            // Clear contents cache
            formula.setLength(0);
            isFormula = true;

            // Decide where to get the formula string from
            String type = attributes.getValue("t");
            if(type != null && type.equals("shared")) {
                // Is it the one that defines the shared, or uses it?
                String ref = attributes.getValue("ref");
                String si = attributes.getValue("si");

                if(ref != null) {
                    // This one defines it
                    // TODO Save it somewhere
                    fIsOpen = true;
                } else {
                    // This one uses a shared formula :(
                }
            } else {
                fIsOpen = true;
            }
        }
        else if("row".equals(localName)) {
            String rowNumStr = attributes.getValue("r");
            previousColumnIndex=0;
            if(rowNumStr != null) {
                rowNum = Integer.parseInt(rowNumStr) - 1;
            } else {
                rowNum = nextRowNum;
            }
        }
        // c => cell
        else if ("c".equals(localName)) {
            // Set up defaults.
            this.nextDataType = xssfDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            cellRef = attributes.getValue("r");
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType))
                nextDataType = xssfDataType.BOOLEAN;
            else if ("e".equals(cellType))
                nextDataType = xssfDataType.ERROR;
            else if ("inlineStr".equals(cellType))
                nextDataType = xssfDataType.INLINE_STRING;
            else if ("s".equals(cellType))
                nextDataType = xssfDataType.SST_STRING;
            else if ("str".equals(cellType))
                nextDataType = xssfDataType.INLINE_STRING;
            else {
                // Number, but almost certainly with a special style or format
                XSSFCellStyle style = null;
                if (stylesTable != null) {
                    if (cellStyleStr != null) {
                        int styleIndex = Integer.parseInt(cellStyleStr);
                        style = stylesTable.getStyleAt(styleIndex);
                    } else if (stylesTable.getNumCellStyles() > 0) {
                        style = stylesTable.getStyleAt(0);
                    }
                }
                if (style != null) {
                    this.formatIndex = style.getDataFormat();
                    this.formatString = style.getDataFormatString();
                    if (this.formatString == null)
                        this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName)
            throws SAXException {

        if (uri != null && ! uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        // v => contents of a cell
        if (isTextTag(localName)) {
            vIsOpen = false;

            if (!isIsOpen) {
                if( hasProcessedHeaders ) {
                    outputCell();
                }
                value.setLength(0);
                isFormula = false;
            }
        } else if ("f".equals(localName)) {
            fIsOpen = false;
        } else if ("is".equals(localName)) {
            isIsOpen = false;
            if( hasProcessedHeaders ) {
                outputCell();
            }
            value.setLength(0);
        } else if ("row".equals(localName)) {
            // Finish up the row

            // some sheets do not have rowNum set in the XML, Excel can read them so we should try to read them as well
            nextRowNum = rowNum + 1;
            hasProcessedHeaders=true;
            System.out.println();

        } else if ("sheetData".equals(localName)) {
            // indicate that this sheet is now done
        }
    }

    /**
     * Captures characters only if a suitable element is open.
     * Originally was just "v"; extended for inlineStr also.
     */
    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
        if (fIsOpen) {
            formula.append(ch, start, length);
        }
    }

    protected void outputCell() {
        final var columnIndex = new CellReference(cellRef).getCol();
        final var expectedType = outputTypes[columnIndex];
        final var outputValue = switch (nextDataType) {
            case NUMBER -> processNumber(expectedType,Double.parseDouble(value.toString()));
            case ERROR -> "**ERR**";
            case FORMULA -> processFormula(expectedType,value.toString());
            case SST_STRING -> processString(expectedType,Short.parseShort(value.toString()));
            case INLINE_STRING -> processString(expectedType,value.toString());
            case BOOLEAN -> processBoolean(expectedType,value.toString());
        };

        for( var i=previousColumnIndex ; i<columnIndex ; i++) {
            System.out.print(',');
        }
        System.out.print(outputValue);
        previousColumnIndex = columnIndex;
    }

    protected String processNumber(ColumnType expected, double value) {

        if(expected==ColumnType.DATETIME && DateUtil.isADateFormat(formatIndex,formatString)) {
            final var tf = DateTimeFormatter.ISO_INSTANT;
            final var tz = TimeZone.getTimeZone("America/Los_Angeles");
            Date d = DateUtil.getJavaDate(value, tz);
            return tf.format(d.toInstant());
        }
        if( expected==ColumnType.BOOLEAN ) {
            return (value==0) ? "FALSE" : "TRUE" ;
        }
        return numberFormat.format(value);
    }
    protected String processString(ColumnType expected, short index) {
        return processString(expected, sharedStringsTable.getItemAt(index).toString());
    }
    protected String processString(ColumnType expected, String value) {
        if(expected==ColumnType.STRING) {
            return value;
        }
        if(expected==ColumnType.DECIMAL) {
            final var num = Double.parseDouble(value);
            return numberFormat.format(num);
        }
        if(expected==ColumnType.BOOLEAN) {
            char c = Character.toUpperCase(value.charAt(0));
            return c=='T'||c=='1'||c=='Y' ? "TRUE" : "FALSE";
        }
        final var dt = DateTimeFormatter.ISO_LOCAL_DATE.parse(value);
        final var tf = DateTimeFormatter.ISO_INSTANT;
        return tf.format(dt);
    }
    protected String processBoolean(ColumnType expected, String value) {
        if(expected==ColumnType.STRING) {
            char c = Character.toUpperCase(value.charAt(0));
            return c=='T'||c=='1'||c=='Y' ? "TRUE" : "FALSE";
        }
        if(expected==ColumnType.DECIMAL) {
            char c = Character.toUpperCase(value.charAt(0));
            return c=='T' ? "1" : "0";
        }
        if(expected==ColumnType.BOOLEAN) {
            char c = Character.toUpperCase(value.charAt(0));
            return c=='T'||c=='1'||c=='Y' ? "TRUE" : "FALSE";
        }
        throw new RuntimeException("Can't convert excel true/false to a date");
    }
    protected String processFormula(ColumnType expected, String value) {
        return value;
    }

}

