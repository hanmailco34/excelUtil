package excel.hanmailco34;

import static excel.hanmailco34.ReflectionUtils.getField;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExcelUtil<T> {

	private SpreadsheetVersion VERSION = SpreadsheetVersion.EXCEL2007;
	private int ROW_INDEX = 0;
	private int COLUMN_INDEX = 0;
	
	private SXSSFWorkbook wb;
	private Sheet sheet;
	private ExcelDto excelDto;
	
	public ExcelUtil(List<T> data, Class<T> type) {
		this.wb = new SXSSFWorkbook();
		this.excelDto = ExcelDtoFactory.mappingExcelDto(type);
		renderExcel(data);
	}
	
	private void renderExcel(List<T> data) {
		sheet = wb.createSheet();
		renderHeaderExcel(sheet, ROW_INDEX++, COLUMN_INDEX);
		
		if (data.isEmpty()) {
			return;
		}
		
		for(Object item : data) {
			renderBodyExcel(item, ROW_INDEX++, COLUMN_INDEX);
		}
	}
	
	private void renderHeaderExcel(Sheet sheet, int rowIndex, int columnIndex) {
		Row row = sheet.createRow(rowIndex);
		int columnStartIndex = columnIndex;
		for(String fieldname : excelDto.getFieldNames()) {
			Cell cell = row.createCell(columnStartIndex++);
			cell.setCellValue(excelDto.getExcelHeaderNames(fieldname));
		}
	}
	
	private void renderBodyExcel(Object item, int rowIndex, int columnIndex) {
		Row row = sheet.createRow(rowIndex);
		int columnStartIndex = columnIndex;
		for(String fieldName : excelDto.getFieldNames()) {			
			try {
				Cell cell = row.createCell(columnStartIndex++);
				Field field = getField(item.getClass(), fieldName);
				field.setAccessible(true);
				Object cellValue = field.get(item);
				renderCellValue(cell, cellValue);
			} catch (Exception e) {
				throw new ExcelException(e.getMessage(), e);
			}
		}
	}
	
	private void renderCellValue(Cell cell, Object cellValue) {
		if(cellValue instanceof Number numberValue) {
			cell.setCellValue(numberValue.doubleValue());
			return;
		}
		cell.setCellValue(cellValue == null ? "" : cellValue.toString());
	}
	
	public void write(OutputStream stream) throws IOException {
		wb.write(stream);
		wb.close();
		wb.dispose();
		stream.close();
	}
	
	public void write(String filePath) throws IOException {
		FileOutputStream fos = new FileOutputStream(new File(filePath));
		wb.write(fos);
      	wb.close();
      	fos.close();
	}
}
