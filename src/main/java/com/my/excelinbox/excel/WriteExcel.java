package com.my.excelinbox.excel;

import io.netty.util.internal.StringUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

import static com.my.excelinbox.excel.ExcelVersion.XLS;
import static com.my.excelinbox.excel.ExcelVersion.XLSX;

/**
 * @author fengran
 */
public class WriteExcel {

    public static @NotNull byte[] write(@NotNull List<?> objects) {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Workbook wb = WriteExcel.setObjects(objects);
            wb.write(out);
            return out.toByteArray();
        } catch (Exception ex) {
            throw new RuntimeException("fail to writing excel:" + ex.getLocalizedMessage(), ex);
        }
    }

    public static @NotNull Workbook setObjects(@NotNull List<?> objects) {
        return setObjects(objects, XLSX);
    }

    public static @NotNull Workbook setObjects(@NotNull List<?> objects, @NotNull ExcelVersion version) {
        if (CollectionUtils.isEmpty(objects)) {
            return emptyWorkBook(version);
        }

        Object object = objects.get(0);
        Class<?> objectClass = object.getClass();
        if (!objectClass.isAnnotationPresent(ExcelSheet.class)) {
            throw new UnsupportedOperationException("Only the class which has annotation @Sheet can be resolve");
        }

        List<Field> fields = Arrays.stream(objectClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .collect(Collectors.toList());
        if (fields.size() == 0) {
            throw new UnsupportedOperationException("Need @ExcelColumn in attribute at least one");
        }

        Workbook workbook = emptyWorkBook(version);
        Sheet sheet = workbook.createSheet();
        Row firstRow = sheet.createRow(0);

        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);

            String headerName = field.getAnnotation(ExcelColumn.class).value();
            if (StringUtil.isNullOrEmpty(headerName)) {
                headerName = field.getName();
            }

            Cell nameCell = firstRow.createCell(i);
            nameCell.setCellValue(headerName);
        }

        for (int j = 1; j <= objects.size(); j++) {
            Row row = sheet.createRow(j);
            try {
                mapObjectToRow(objects.get(j - 1), row, fields);
            } catch (Exception ex) {
                RuntimeException excelException = new RuntimeException("Error in mapping excel: " + ex.getMessage());
                excelException.setStackTrace(ex.getStackTrace());
                throw excelException;
            }

        }

        return workbook;
    }

    private static void mapObjectToRow(@NotNull Object object, Row row, List<Field> fields) throws Exception {
        for (int i = 0; i < fields.size(); i++) {
            Cell cell = row.createCell(i);

            Field field = fields.get(i);
            field.setAccessible(true);
            Object value = field.get(object);
            if (value == null) {
                cell.setCellValue("");
            } else {
                if (Date.class.equals(field.getType())) {
                    cell.setCellValue((Date) value);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
    }

    private static @NotNull Workbook emptyWorkBook(@NotNull ExcelVersion version) {
        if (XLS.equals(version)) {
            return new HSSFWorkbook();
        } else {
            return new XSSFWorkbook();
        }
    }
}
