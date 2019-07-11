package com.my.excelinbox.excel;

import io.netty.util.internal.StringUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.function.Predicate;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * @author fengran
 */
public class ReadExcel {

    public static @NotNull <T> List<T> getObjectsFromXLS(InputStream is, Class<T> objectClass) {
        try (InputStream excelIs = is){
            return getObjects(new HSSFWorkbook(excelIs), objectClass, 0);
        }catch (Exception ex){
            throw new RuntimeException(ex);
        }
    }

    public static @NotNull <T> List<T> getObjectsFromXLSX(InputStream is, Class<T> objectClass) {
        try (InputStream excelIs = is){
            return getObjects(new XSSFWorkbook(excelIs), objectClass, 0);
        }catch (Exception ex){
            throw new RuntimeException(ex);
        }
    }

    public static @NotNull <T> List<T> getObjects(Workbook workbook, Class<T> objectClass) {
        return getObjects(workbook, objectClass, 0);
    }

    public static @NotNull <T> List<T> getObjects(Workbook workbook, Class<T> objectClass, Integer sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        return getObjects(sheet, objectClass);
    }

    public static @NotNull <T> List<T> getObjects(Workbook workbook, Class<T> objectClass, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        return getObjects(sheet, objectClass);
    }

    private static @NotNull <T> List<T> getObjects(Sheet sheet, Class<T> objectClass) {
        List<T> result = new LinkedList<>();

        int cowNum = sheet.getLastRowNum();
        if (cowNum == 0) {
            return result;
        }

        Row preHeader = sheet.getRow(0);
        Map<Integer, String> headerMap = new HashMap<>();
        for (int i = 0; i <= preHeader.getLastCellNum(); i++) {
            String headerName;

            try {
                headerName = preHeader.getCell(i).getStringCellValue();
            } catch (NullPointerException ex) {
                break;
            }
            if (StringUtil.isNullOrEmpty(headerName)) {
                break;
            }

            headerMap.put(i, headerName);
        }
        SheetHeader header = new SheetHeader(headerMap);

        try {
            //sheet.getLastRowNum()可能大于实际行数
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (Objects.isNull(row)) {
                    break;
                }
                T object = mapRowToObject(objectClass, row, header);
                //存在空白行则直接结束
                if (object == null) {
                    break;
                }
                result.add(object);
            }
            return result;
        } catch (Exception ex) {
            RuntimeException excelException = new RuntimeException("Error in mapping excel: " + ex.getMessage());
            excelException.setStackTrace(ex.getStackTrace());
            throw excelException;
        }
    }

    private static <T> T mapRowToObject(Class<T> objectClass, @NotNull Row row, SheetHeader sheetHeader) throws Exception {
        if (!objectClass.isAnnotationPresent(ExcelSheet.class)) {
            throw new UnsupportedOperationException("Only the class which has annotation @Sheet can be resolve");
        }

        //<Excel列名, 实体属性>, 若列名未填写则使用属性名
        HashMap<String, Field> fieldMap = new HashMap<>(10);
        Arrays.stream(objectClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .peek(field -> field.setAccessible(true))
                .forEach(field -> {
                    String annotationName = field.getAnnotation(ExcelColumn.class).value();
                    if (StringUtil.isNullOrEmpty(annotationName)) {
                        fieldMap.put(field.getName(), field);
                    } else {
                        fieldMap.put(annotationName, field);
                    }
                });

        if (fieldMap.size() == 0) {
            throw new UnsupportedOperationException("Need @ExcelColumn in attribute at least one");
        }

        T object = objectClass.newInstance();

        // 有效的属性数
        // {null,"", 0, 0.00} 均会被认为无效值
        int goodAttributeNum = 0;
        for (int i = 0; i < sheetHeader.size(); i++) {
            String columnName = sheetHeader.getColumnName(i);

            if (!fieldMap.containsKey(columnName)) {
                continue;
            }
            Field field = fieldMap.get(columnName);

            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            goodAttributeNum = setObjectAttribute(field, cell, object, goodAttributeNum);
        }

        // 有效的属性数为0时将会认为该对象无效，强制返回null
        if (goodAttributeNum == 0) {
            return null;
        } else {
            return object;
        }
    }

    private static Predicate<Class> isInt = c -> int.class.equals(c) || Integer.class.equals(c);
    private static Predicate<Class> isShort = c -> short.class.equals(c) || Short.class.equals(c);
    private static Predicate<Class> isLong = c -> long.class.equals(c) || Long.class.equals(c);
    private static Predicate<Class> isFloat = c -> float.class.equals(c) || Float.class.equals(c);
    private static Predicate<Class> isDouble = c -> double.class.equals(c) || Double.class.equals(c);

    private static <T> int setObjectAttribute(Field field, Cell cell, T object, Integer goodAttributeNum) throws Exception {
        Class fieldClass = field.getType();
        field.setAccessible(true);

        if (cell.getCellType() == _NONE || cell.getCellType() == BLANK) {
            return goodAttributeNum;
        }

        if (cell.getCellType() == STRING) {
            String preValue = cell.getStringCellValue();
            if (preValue == null || "".equals(preValue)) {
                return goodAttributeNum;
            } else {
                goodAttributeNum++;
            }

            if (String.class.equals(fieldClass)) {
                field.set(object, preValue);
            } else if (Date.class.equals(fieldClass)) {
                field.set(object, cell.getDateCellValue());
            } else if (isInt.test(fieldClass)) {
                Integer value = Integer.valueOf(preValue);
                field.set(object, value);
            } else if (isShort.test(fieldClass)) {
                Short value = Short.valueOf(preValue);
                field.set(object, value);
            } else if (isLong.test(fieldClass)) {
                Long value = Long.valueOf(preValue);
                field.set(object, value);
            } else if (isFloat.test(fieldClass)) {
                Float value = Float.valueOf(preValue);
                field.set(object, value);
            } else if (isDouble.test(fieldClass)) {
                Double value = Double.valueOf(preValue);
                field.set(object, value);
            }
        }

        if (cell.getCellType() == NUMERIC) {
            Double preValue = cell.getNumericCellValue();
            if (!preValue.equals(Double.NaN)) {
                goodAttributeNum++;
            }

            if (String.class.equals(fieldClass)) {
                String value = String.valueOf(preValue.intValue());
                field.set(object, value);
            } else if (isInt.test(fieldClass)) {
                Integer value = preValue.intValue();
                field.set(object, value);
            } else if (isShort.test(fieldClass)) {
                Short value = preValue.shortValue();
                field.set(object, value);
            } else if (isLong.test(fieldClass)) {
                Long value = preValue.longValue();
                field.set(object, value);
            } else if (isFloat.test(fieldClass)) {
                Float value = preValue.floatValue();
                field.set(object, value);
            } else if (isDouble.test(fieldClass)) {
                field.set(object, preValue);
            }
        }

        return goodAttributeNum;
    }
}
