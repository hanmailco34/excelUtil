package excel.hanmailco34;

import static excel.hanmailco34.ReflectionUtils.getAllFields;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelDtoFactory {

	public static ExcelDto mappingExcelDto(Class<?> type) {
		Map<String, String> headerNames = new LinkedHashMap<>();
		List<String> fieldNames = new ArrayList<>();
		Class<ExcelColumn> excelClass = ExcelColumn.class;
		
		for(Field field : getAllFields(type)) {
			if(field.isAnnotationPresent(excelClass)) {
				ExcelColumn annoExcelColumn = field.getAnnotation(excelClass);
				fieldNames.add(field.getName());
				headerNames.put(field.getName(), annoExcelColumn.headerName());
			}
		}
		
		return new ExcelDto(headerNames, fieldNames);		
	}
}
