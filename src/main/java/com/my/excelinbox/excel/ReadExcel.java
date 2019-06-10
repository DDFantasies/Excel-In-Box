package com.my.excelinbox.excel;

import io.netty.util.internal.StringUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ReadExcel {

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

    public static @NotNull <T> List<T> getObjects(Sheet sheet, Class<T> objectClass) {
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
            //sheet.getLastRowNum()可能大于实际行数，以ExcelId为准
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (Objects.isNull(row)) {
                    break;
                }
                T object = mapRowToObject(objectClass, row, header);
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
        HashMap<String, Field> fieldMap = new HashMap<>();
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
        AtomicInteger goodAttributeNum = new AtomicInteger(0);

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
            String personalDateFormat = field.getAnnotation(ExcelColumn.class).dateFormat();
            setObjectAttribute(field, cell, object, personalDateFormat, goodAttributeNum);
        }

        // 有效的属性数为0时将会认为该对象无效，强制返回null
        if (goodAttributeNum.get() == 0) {
            return null;
        } else {
            return object;
        }
    }

    private static <T> void setObjectAttribute(Field field, Cell cell, T object, String dateFormat, AtomicInteger goodAttributeNum) throws Exception {
        Class fieldClass = field.getType();
        field.setAccessible(true);

        //这里写的好垃圾啊，然而没想到怎么改进
        boolean isInt = int.class.equals(fieldClass) || Integer.class.equals(fieldClass);
        boolean isShort = short.class.equals(fieldClass) || Short.class.equals(fieldClass);
        boolean isLong = long.class.equals(fieldClass) || Long.class.equals(fieldClass);
        boolean isFloat = float.class.equals(fieldClass) || Float.class.equals(fieldClass);
        boolean isDouble = double.class.equals(fieldClass) || Double.class.equals(fieldClass);

        if (cell.getCellType() == _NONE || cell.getCellType() == BLANK) {
            return;
        }

        if (cell.getCellType() == STRING) {
            String preValue = cell.getStringCellValue();
            if (preValue == null || preValue.equals("")) {
                return;
            } else {
                goodAttributeNum.incrementAndGet();
            }

            if (String.class.equals(fieldClass)) {
                field.set(object, preValue);
            } else if (Date.class.equals(fieldClass)) {
                Date value = new SimpleDateFormat(dateFormat).parse(preValue);
                field.set(object, value);
            } else if (isInt) {
                Integer value = Integer.valueOf(preValue);
                field.set(object, value);
            } else if (isShort) {
                Short value = Short.valueOf(preValue);
                field.set(object, value);
            } else if (isLong) {
                Long value = Long.valueOf(preValue);
                field.set(object, value);
            } else if (isFloat) {
                Float value = Float.valueOf(preValue);
                field.set(object, value);
            } else if (isDouble) {
                Double value = Double.valueOf(preValue);
                field.set(object, value);
            }
        }

        if (cell.getCellType() == NUMERIC) {
            Double preValue = cell.getNumericCellValue();
            if (!preValue.equals(0.00)) {
                goodAttributeNum.incrementAndGet();
            }

            if (String.class.equals(fieldClass)) {
                String value = String.valueOf(preValue.intValue());
                field.set(object, value);
            } else if (isInt) {
                Integer value = preValue.intValue();
                field.set(object, value);
            } else if (isShort) {
                Short value = preValue.shortValue();
                field.set(object, value);
            } else if (isLong) {
                Long value = preValue.longValue();
                field.set(object, value);
            } else if (isFloat) {
                Float value = preValue.floatValue();
                field.set(object, value);
            } else if (isDouble) {
                field.set(object, preValue);
            }
        }

    }
}
