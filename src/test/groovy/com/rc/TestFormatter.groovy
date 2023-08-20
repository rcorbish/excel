package com.rc

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.xssf.streaming.SXSSFCell
import spock.lang.Specification

class TestFormatter extends Specification {
	char[] columnTypes = ['c' as char]
	XlsxDataFormatter instance

	def setup() {
		instance = new XlsxDataFormatter(columnTypes)
	}


	def "Should format a number nicely"() {

		def cell1 = Mock(Cell)

		cell1.getCellType() >> celltype
		cell1.getNumericCellValue() >> number

		def ans = instance.formatCellValue(cell1)
	expect:
		ans == expected

	where:
		celltype 			| number 		| string 					| expected
		CellType.NUMERIC 	| 12345.6789 	| null 						| "12345.6789"
	}
}
