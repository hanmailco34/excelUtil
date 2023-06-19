package excel.hanmailco34;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

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
	
	public static <T> T mapToClass(Map<String, Object> map, Class<T> type) {
		T instance = null;
		try {
			instance = type.getConstructor().newInstance();
			for(Map.Entry<String, Object> entry : map.entrySet()) {
				Field[] fields = type.getDeclaredFields();
				
				for(Field field : fields) {
					field.setAccessible(true);
					
					String fieldName = field.getName();
					
					if(entry.getValue() == null) continue;
					
					boolean isSameType = entry.getValue().getClass().equals(field.getType());
					boolean isSameName = entry.getKey().equals(fieldName);
					
					if(isSameType && isSameName) {
						field.set(instance, map.get(fieldName));
					}
				}
			}
		} catch (InstantiationException | IllegalAccessException | IllegalArgumentException | InvocationTargetException
				| NoSuchMethodException | SecurityException e) {
			e.printStackTrace();
		}
		
		return instance;
	}
}
