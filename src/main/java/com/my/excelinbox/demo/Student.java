package com.my.excelinbox.demo;

import com.my.excelinbox.excel.ExcelColumn;
import com.my.excelinbox.excel.ExcelSheet;

import java.util.Date;

@ExcelSheet
public class Student {

    @ExcelColumn
    private String no;

    @ExcelColumn("名字")
    private String name;

    @ExcelColumn("年龄")
    private int age;

    @ExcelColumn("出生日期")
    private Date birthday;

    @ExcelColumn("学费")
    private double fee;

    public String getNo() {
        return no;
    }

    public void setNo(String no) {
        this.no = no;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public double getFee() {
        return fee;
    }

    public void setFee(double fee) {
        this.fee = fee;
    }
}
