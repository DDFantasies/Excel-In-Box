package com.my.excelinbox.excel;

import java.lang.annotation.Documented;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * @author fengran
 */
@Documented
@Target(TYPE)
@Retention(RUNTIME)
@Inherited
//标记excel实体类
public @interface ExcelSheet {
}
