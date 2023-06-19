package excel.hanmailco34;

import static excel.hanmailco34.ReflectionUtils.getField;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil<T> {

	private SpreadsheetVersion VERSION = SpreadsheetVersion.EXCEL2007;
	private int ROW_INDEX = 0;
	private int COLUMN_INDEX = 0;
	
	private SXSSFWorkbook wb;
	private XSSFWorkbook xb;
	private Sheet sheet;
	private ExcelDto excelDto;
	private Class<T> type;	
	
	public ExcelUtil(List<T> data, Class<T> type) {
		this.wb = new SXSSFWorkbook();
		this.excelDto = ExcelDtoFactory.mappingExcelDto(type);
		renderExcel(data);
	}
	
	public ExcelUtil(File file, Class<T> type) {
		FileInputStream fis;
		try {
			fis = new FileInputStream(file);
			this.xb = new XSSFWorkbook(fis);
			this.excelDto = ExcelDtoFactory.mappingExcelDto(type);
			this.type = type;
		} catch (IOException e) {
			e.printStackTrace();
		}		
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
	
	private List<String> getExcelTile() {
		Iterator<Cell> iterList = sheet.getRow(0).iterator();
		List<String> titles = new ArrayList<>();
		while(iterList.hasNext()) {
			titles.add(String.valueOf(iterList.next()));
		}
		return titles;
	}
	
	private Map<Integer, String> getCellIndex(List<String> titles) {
		Map<String, String> excelHeaders = this.excelDto.getExcelHeaderNames();
		Map<Integer, String> result = new HashMap<>();
		
		for(Map.Entry<String, String> entry : excelHeaders.entrySet()) {
			String value = entry.getValue();
			String key = entry.getKey();
			for(int index = 0; index < titles.size(); index++) {
				String title = titles.get(index);
				if(value.equals(title)) {
					result.put(index, key);
					break;
				}
			}
		}
			
		return result;
	}
	
	private T readRowData(Row row, Map<Integer, String> cells) {
		Map<String, Object> rowData = new HashMap<>();
		
		for(Map.Entry<Integer, String> entry : cells.entrySet()) {
			int cellNo = entry.getKey();
			String columnName = entry.getValue();
			Cell cell = row.getCell(cellNo);
			String value = "";
			
			if(cell != null) {
				switch(cell.getCellType()) {
	                case NUMERIC:
	                    value = cell.getNumericCellValue() + "";
	                    break;
	                case STRING:
	                    value = cell.getStringCellValue() + "";
	                    break;
                }
			}
			rowData.put(columnName, value);
		}
		return ReflectionUtils.mapToClass(rowData, type);
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
	
	public List<T> convertToList() {
		this.sheet = xb.getSheetAt(0);
		int rowNo = 0;
        int rows = sheet.getPhysicalNumberOfRows();
        
        List<String> titles = getExcelTile();
        Map<Integer, String> cells = getCellIndex(titles);
        
        List<T> result = new ArrayList<>();
    
		for(rowNo = 1; rowNo < rows; rowNo++){
			Row row = sheet.getRow(rowNo);
			result.add(readRowData(row, cells));
		}
		
		return result;
	}
}
