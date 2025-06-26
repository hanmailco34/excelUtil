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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil<T> {

	private int ROW_INDEX = 0;
	private int COLUMN_INDEX = 0;
	private int SHEET_INDEX = 0;
	private String SHEET_NAME = "";
	
	private SXSSFWorkbook wb;
	private XSSFWorkbook xb;
	private Sheet sheet;
	private ExcelDto excelDto;
	private Class<T> type;
	
	/**
	 * 리스트 -> 엑셀로 변환할때 쓰는 생성자
	 * @return
	 */
	public ExcelUtil(List<T> data, Class<T> type, String sheetName) {
		if(sheetName == null || sheetName.trim().isEmpty())
			this.SHEET_NAME = null;
		else
			this.SHEET_NAME = sheetName;
		this.wb = new SXSSFWorkbook(data.size()+1);
		this.excelDto = ExcelDtoFactory.mappingExcelDto(type);
		renderExcel(data);
	}
	
	public ExcelUtil(List<T> data, Class<T> type) {
		this(data, type, null);
	}
	
	/**
	 * 엑셀파일 -> 리스트로 변활할때 쓰는 생성자
	 * @param file
	 * @param type
	 */
	public ExcelUtil(File file, Class<T> type) {
		try(FileInputStream fis = new FileInputStream(file)) {
			this.xb = new XSSFWorkbook(fis);
			this.excelDto = ExcelDtoFactory.mappingExcelDto(type);
			this.type = type;
		} catch (IOException e) {
			e.printStackTrace();
		}		
	}
	
	public ExcelUtil(File file, Class<T> type, int sheetIndex) {
		this(file, type);
		this.SHEET_INDEX = sheetIndex;
	}	
	
	/**
	 * 하나의 엑셀 파일에 여러개의 시트
	 * @param list
	 */
	public ExcelUtil(List<ExcelUtil<?>> list) {
		this.wb = new SXSSFWorkbook();
		for(ExcelUtil<?> item : list) {
			Sheet sourceSheet = item.sheet;
			setSheetName(item.SHEET_NAME);
			copyRows(sourceSheet);
		}
	}
	
	private void setSheetName(String sheetName) {
		this.sheet = sheetName == null ? wb.createSheet() : wb.createSheet(sheetName);
	}
	
	private void renderExcel(List<T> data) {
		setSheetName(this.SHEET_NAME);
		
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
		if(cellValue instanceof Number) {
			Number numberValue = (Number) cellValue;
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
			Object value = "";
			if(cell != null) {
				switch(cell.getCellType()) {
	                case FORMULA:
	                    value = cell.getCellFormula();
	                    break;
	                case NUMERIC:
	                    value = cell.getNumericCellValue();
	                    break;
	                case STRING:
	                    value = cell.getStringCellValue();
	                    break;
					case BOOLEAN:
						value = cell.getBooleanCellValue() + "";
						break;
	                case ERROR:
	                    value = cell.getErrorCellValue() + "";
	                    break;
					case BLANK:
						break;
                }
			}
			rowData.put(columnName, value);
		}
		return ReflectionUtils.mapToClass(rowData, type);
	}
	
	private void copyRows(Sheet sourceSheet) {
        for (Row sourceRow : sourceSheet) {
            Row targetRow = this.sheet.createRow(sourceRow.getRowNum());

            for (Cell cell : sourceRow) {
                Cell targetCell = targetRow.createCell(cell.getColumnIndex(), cell.getCellType());
                targetCell.setCellStyle(cell.getCellStyle());

                switch (cell.getCellType()) {
	                case FORMULA:
	                	targetCell.setCellFormula(cell.getCellFormula());
	                    break;
	                case NUMERIC:
	                	targetCell.setCellValue(cell.getNumericCellValue());
	                    break;
	                case STRING:
	                	targetCell.setCellValue(cell.getStringCellValue());
	                    break;
	                case BLANK:
	                    break;
	                case ERROR:
	                	targetCell.setCellErrorValue(cell.getErrorCellValue());
	                    break;
                }
            }
        }
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
		this.sheet = xb.getSheetAt(this.SHEET_INDEX);
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
