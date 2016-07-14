package org.javafunk.excelparser;

import org.javafunk.excelparser.annotations.ExcelField;
import org.javafunk.excelparser.annotations.ExcelObject;
import org.javafunk.excelparser.annotations.MappedExcelObject;
import org.javafunk.excelparser.annotations.ParseType;
import org.javafunk.excelparser.exception.ExcelInvalidCell;
import org.javafunk.excelparser.exception.ExcelInvalidCellValuesException;
import org.javafunk.excelparser.exception.ExcelParsingException;
import org.javafunk.excelparser.helper.HSSFHelper;
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
import java.util.OptionalInt;
import java.util.Set;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class SheetParser {
    List<ExcelInvalidCell> excelInvalidCells;

    public SheetParser()
    {
        excelInvalidCells = new ArrayList<>();
    }

    public <T> List<T> createEntity(Sheet sheet, Class<T> clazz, Consumer<ExcelParsingException> errorHandler) {
        List<T> list = new ArrayList<>();
        ExcelObject excelObject = getExcelObject(clazz, errorHandler);
        if (excelObject.start() <= 0 || excelObject.end() < 0) {
            return list;
        }
        int end = getEnd(sheet, clazz, excelObject);

        for (int currentLocation = excelObject.start(); currentLocation <= end; currentLocation++) {
            T object = getNewInstance(sheet, clazz, excelObject.parseType(), currentLocation, excelObject.zeroIfNull(),
                    errorHandler);
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

    private <T> int getEnd(Sheet sheet, Class<T> clazz, ExcelObject excelObject) {
        int end = excelObject.end();
        if (end > 0) {
            return end;
        }
        return getRowOrColumnEnd(sheet, clazz);
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

    public <T> int getRowOrColumnEnd(Sheet sheet, Class<T> clazz) {
        ExcelObject excelObject = getExcelObject(clazz, e -> {
            throw e;
        });
        ParseType parseType = excelObject.parseType();
        if (parseType == ParseType.ROW) {
            return sheet.getLastRowNum() + 1;
        }

        Set<Integer> positions = getExcelFieldPositionMap(clazz).keySet();
        OptionalInt maxPosition = positions.stream().mapToInt((x) -> x).max();
        OptionalInt minPosition = positions.stream().mapToInt((x) -> x).min();

        int maxCellNumber = 0;
        for (int i = minPosition.getAsInt(); i < maxPosition.getAsInt(); i++) {
            int cellsNumber = sheet.getRow(i).getLastCellNum();
            if (maxCellNumber < cellsNumber) {
                maxCellNumber = cellsNumber;
            }
        }
        return maxCellNumber;
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
            Object cellValueString;
            if (ParseType.ROW == parseType) {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), currentLocation, position, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(sheet, String.class, currentLocation, position, zeroIfNull, errorHandler);
            } else {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), position, currentLocation, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(sheet, String.class, position, currentLocation, zeroIfNull, errorHandler);
            }
            ExcelField annotation = field.getAnnotation(ExcelField.class);
            if (annotation.validate())
            {
                Pattern pattern =  Pattern.compile(annotation.regex());
                cellValueString = cellValueString != null ? cellValueString.toString() : "";
                Matcher matcher = pattern.matcher((String)cellValueString);
                if(!matcher.matches())
                {
                    ExcelInvalidCell excelInvalidCell = new ExcelInvalidCell(position,currentLocation, (String)cellValueString);
                    excelInvalidCells.add(excelInvalidCell);
                    if(annotation.validationType() == ExcelField.ValidationType.HARD)
                    {
                        throw new ExcelInvalidCellValuesException("Invalid cell value at [" + currentLocation + ", " + position + "] in the sheet. This exception can be suppressed by setting 'validationType' in @ExcelField to 'ValidationType.SOFT");
                    }
                }
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
