package com.my.excelinbox.excel;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

@Target({FIELD})
@Retention(RUNTIME)
//标记实体属性对应的的excel列名
public @interface ExcelColumn {
    //列名，未填写则使用属性名，同一excel文件中不可重复出现
    String value() default "";
    String dateFormat() default "yyyy/MM/dd";
}
