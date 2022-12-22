package excel.hanmailco34;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ReflectionUtils {
	
	private ReflectionUtils() {
		
	}
	
	public static Field getField(Class<?> type, String name) throws NoSuchFieldException {
		return type.getDeclaredField(name);
	}

	public static List<Field> getAllFields(Class<?> type) {
		List<Field> fields = new ArrayList<>();
		fields.addAll(Arrays.asList(type.getDeclaredFields()));
		return fields;
	}
}
