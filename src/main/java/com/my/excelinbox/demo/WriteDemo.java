package com.my.excelinbox.demo;

import com.my.excelinbox.excel.WriteExcel;

import java.io.FileOutputStream;

import java.util.Collections;
import java.util.Date;

/**
 * @author fengran
 */
public class WriteDemo {
    public static void main(String[] args) {
        Student student = new Student();
        student.setName("人造人0号");
        student.setNo("2019040010");
        student.setAge(20);
        student.setBirthday(new Date());
        student.setFee(1500.58);

        byte[] result = WriteExcel.write(Collections.singletonList(student));
        try(FileOutputStream fos = new FileOutputStream("/test.xlsx")){
            fos.write(result);
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }
}
