package excel.hanmailco34;

import java.util.List;
import java.util.Map;

public class ExcelDto {

	private Map<String, String> excelHeaderNames;
	private List<String> fieldNames;
	
	public ExcelDto(Map<String, String> excelHeaderNames, List<String> fieldNames) {
		this.excelHeaderNames = excelHeaderNames;
		this.fieldNames = fieldNames;
	}
	
	public String getExcelHeaderNames(String fieldname) {
		return excelHeaderNames.get(fieldname);
	}
	public List<String> getFieldNames() {
		return fieldNames;
	}
}
