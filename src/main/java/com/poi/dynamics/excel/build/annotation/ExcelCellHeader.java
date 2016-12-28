package com.poi.dynamics.excel.build.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.TYPE;

/**
 * Created by lfwang on 2016/12/28.
 */
@Documented
@Target({TYPE, FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellHeader {

    String value() default "";
}
