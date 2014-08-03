package com.thoughtworks.excelparser;

import com.thoughtworks.excelparser.annotations.ExcelField;
import com.thoughtworks.excelparser.annotations.ExcelObject;
import com.thoughtworks.excelparser.annotations.MappedExcelObject;
import com.thoughtworks.excelparser.annotations.ParseType;
import com.thoughtworks.excelparser.exception.ExcelParsingException;
import com.thoughtworks.excelparser.helper.HSSFHelper;
import lombok.AccessLevel;
import lombok.experimental.FieldDefaults;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class SheetParser {

    public <T> List<T> createEntity(Sheet sheet, Class<T> clazz, Consumer<ExcelParsingException> errorHandler) {
        List<T> list = new ArrayList<>();
        ExcelObject excelObject = getExcelObject(clazz, errorHandler);
        for (int currentLocation = excelObject.start(); currentLocation <= excelObject.end(); currentLocation++) {
            T object = getNewInstance(sheet, clazz, excelObject.parseType(), currentLocation, excelObject.zeroIfNull(), errorHandler);
            List<Field> mappedExcelFields = getMappedExcelObjects(clazz);
            for (Field mappedField : mappedExcelFields) {
                Class<?> fieldType = mappedField.getType();
                Class<?> clazz1 = fieldType.equals(List.class) ? getFieldType(mappedField) : fieldType;
                List<?> fieldValue = createEntity(sheet, clazz1, errorHandler);
                if (fieldType.equals(List.class)) {
                    setFieldValue(mappedField, object, fieldValue);
                } else if (!fieldValue.isEmpty()) {
                    setFieldValue(mappedField, object, fieldValue.get(0));
                }
            }
            list.add(object);
        }
        return list;
    }

    /**
     * @deprecated Pass an error handler lambda instead (see other signature)
     */
    @Deprecated
    public <T> List<T> createEntity(Sheet sheet, String sheetName, Class<T> clazz) {
        return createEntity(sheet, clazz, error -> {
            throw error;
        });
    }

    private Class<?> getFieldType(Field field) {
        Type type = field.getGenericType();
        if (type instanceof ParameterizedType) {
            ParameterizedType pt = (ParameterizedType) type;
            return (Class<?>) pt.getActualTypeArguments()[0];
        }

        return null;
    }

    private <T> List<Field> getMappedExcelObjects(Class<T> clazz) {
        List<Field> fieldList = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            MappedExcelObject mappedExcelObject = field.getAnnotation(MappedExcelObject.class);
            if (mappedExcelObject != null) {
                field.setAccessible(true);
                fieldList.add(field);
            }
        }
        return fieldList;
    }

    private <T> ExcelObject getExcelObject(Class<T> clazz, Consumer<ExcelParsingException> errorHandler) {
        ExcelObject excelObject = clazz.getAnnotation(ExcelObject.class);
        if (excelObject == null) {
            errorHandler.accept(new ExcelParsingException("Invalid class configuration - ExcelObject annotation missing - " + clazz.getSimpleName()));
        }
        return excelObject;
    }

    private <T> T getNewInstance(Sheet sheet, Class<T> clazz, ParseType parseType, Integer currentLocation, boolean zeroIfNull, Consumer<ExcelParsingException> errorHandler) {
        T object = getInstance(clazz, errorHandler);
        Map<Integer, Field> excelPositionMap = getExcelFieldPositionMap(clazz);
        for (Integer position : excelPositionMap.keySet()) {
            Field field = excelPositionMap.get(position);
            Object cellValue;
            if (ParseType.ROW == parseType) {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), currentLocation, position, zeroIfNull, errorHandler);
            } else {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), position, currentLocation, zeroIfNull, errorHandler);
            }
            setFieldValue(field, object, cellValue);
        }

        return object;
    }

    private <T> T getInstance(Class<T> clazz, Consumer<ExcelParsingException> errorHandler) {
        T object;
        try {
            Constructor<T> constructor = clazz.getDeclaredConstructor();
            constructor.setAccessible(true);
            object = constructor.newInstance();
        } catch (Exception e) {
            errorHandler.accept(new ExcelParsingException("Exception occurred while instantiating the class " + clazz.getName(), e));
            return null;
        }
        return object;
    }

    private <T> void setFieldValue(Field field, T object, Object cellValue) {
        try {
            field.set(object, cellValue);
        } catch (IllegalArgumentException | IllegalAccessException e) {
            throw new ExcelParsingException("Exception occurred while setting field value ", e);
        }
    }

    private <T> Map<Integer, Field> getExcelFieldPositionMap(Class<T> clazz) {
        Map<Integer, Field> fieldMap = new HashMap<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null) {
                field.setAccessible(true);
                fieldMap.put(excelField.position(), field);
            }
        }
        return fieldMap;
    }

}
