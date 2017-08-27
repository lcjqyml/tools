package com.milin.tools.excel;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public final class ExcelWriter {
	public static final class Cell {
		private Workbook wb;
		private org.apache.poi.ss.usermodel.Cell target;

		private Cell(Workbook wb, org.apache.poi.ss.usermodel.Cell cell) {
			this.wb = wb;
			this.target = cell;
		}

		public void setCellValue(String value) {
			this.target.setCellValue(value);
		}

		public void setCellValue(Date value) {
			if (null == value) {
				this.setCellValue("未设置");
				return;
			}
			CellStyle style = this.wb.createCellStyle();
			DataFormat format = wb.createDataFormat();
			style.setDataFormat(format.getFormat("yyyy/m/d hh:mm:ss"));
			this.target.setCellStyle(style);
			this.target.setCellValue(value);
		}

		public void setCellValue(double value) {
			CellStyle style = this.wb.createCellStyle();
			DataFormat format = wb.createDataFormat();
			style.setDataFormat(format.getFormat("###,##0"));
			this.target.setCellStyle(style);
			this.target.setCellValue(value);
		}

		public void setCellStyle(CellStyle style) {
			this.target.setCellStyle(style);
		}
	}

	public static final class Sheet {
		private Workbook wb;
		private org.apache.poi.ss.usermodel.Sheet sheet;

		private Sheet(Workbook wb, org.apache.poi.ss.usermodel.Sheet sheet) {
			this.sheet = sheet;
			this.wb = wb;
		}

		public Cell getOrCreateCell(int rowBasedZero, int cellBasedZero) {
			Row row = this.sheet.getRow(rowBasedZero);
			if (null == row) {
				row = this.sheet.createRow(rowBasedZero);
			}
			return new Cell(this.wb, row.getCell(cellBasedZero,
					Row.CREATE_NULL_AS_BLANK));
		}

		// 合并单元格
		public void mergedRegion(int intFirstRow, int intLastRow,
				int intFirstColumn, int intLastColumn) {
			this.sheet.addMergedRegion(new CellRangeAddress(intFirstRow,
					intLastRow, intFirstColumn, intLastColumn));
		}

		// 添加居中的样式
		public CellStyle addStyle() {
			CellStyle style = wb.createCellStyle();
			style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平
			return style;
		}

	}

	private Map<String, Sheet> sheetMap = new HashMap<String, Sheet>();
	private OutputStream out;
	private Workbook wb;

	public ExcelWriter(OutputStream out) {
		this.out = out;
		wb = new XSSFWorkbook();
	}

	public Sheet getOrCreateSheet(String sheet) {
		Sheet sheet1 = this.sheetMap.get(sheet);
		if (null != sheet1) {
			return sheet1;
		}
		org.apache.poi.ss.usermodel.Sheet xlsSheet = this.wb.getSheet(sheet);
		if (null == xlsSheet) {
			xlsSheet = this.wb.createSheet(sheet);
		}
		Sheet obj = new Sheet(this.wb, xlsSheet);
		this.sheetMap.put(sheet, obj);
		return obj;
	}

	public void write() throws IOException {
		this.wb.write(this.out);
	}

	public void close() {
		this.sheetMap.clear();
		IOUtils.closeQuietly(this.out);
	}
}