package com.my.excelinbox.service;

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

    @ExcelColumn(value = "登记时间", dateFormat = "yyyy-MM-dd HH:mm:ss")
    private Date registerTime;

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

    public Date getRegisterTime() {
        return registerTime;
    }

    public void setRegisterTime(Date registerTime) {
        this.registerTime = registerTime;
    }

    public double getFee() {
        return fee;
    }

    public void setFee(double fee) {
        this.fee = fee;
    }
}
