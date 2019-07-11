package com.my.excelinbox.demo;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.my.excelinbox.excel.ReadExcel;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.List;

/**
 * @author fengran
 */
public class ReadDemo {
    public static void main(String[] args) {
        try(InputStream is = ReadDemo.class.getResourceAsStream("/application.yml")){
            List<Student> students = ReadExcel.getObjectsFromXLSX(is, Student.class);
            ObjectMapper mapper = new ObjectMapper();
            System.out.println(mapper.writeValueAsString(students));
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }
}
