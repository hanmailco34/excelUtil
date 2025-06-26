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
	
	public static Field getField(Class<?> type, String fieldName) {
		try {
			return type.getDeclaredField(fieldName);
		} catch (NoSuchFieldException | SecurityException e) {
			Class<?> superClass = type.getSuperclass();
			if (superClass != null) {
	            return getField(superClass, fieldName);
	        }
	        return null;
		}
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
				String fieldName = entry.getKey();
	            Object value = entry.getValue();
	            
	            if(value == null) continue;
	            
	            Field field = getField(type, fieldName);
	            
	            if(field != null && field.getType().equals(value.getClass())) {
	                field.setAccessible(true);
	                field.set(instance, value);
	            }else if(field != null && field.getType().equals(String.class) && value instanceof Number){ // 명시적 타입 컨버팅 (Number -> String)
					if(value.getClass().equals(Double.class) && ((Double) value).doubleValue() == ((Double) value).intValue()){
						value =  Integer.valueOf(((Double) value).intValue());
					}
					field.setAccessible(true);
					field.set(instance, value+"");
				}
			}
		} catch (InstantiationException | IllegalAccessException | IllegalArgumentException | InvocationTargetException
				| NoSuchMethodException | SecurityException e) {
			e.printStackTrace();
		}
		
		return instance;
	}
}
