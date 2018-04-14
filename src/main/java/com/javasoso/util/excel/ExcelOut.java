package com.javasoso.util.excel;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出
 *
 * @author jasonzhu
 * @date 2018/4/14
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelOut {
    /**
     * 列数 0 第一列
     */
    int value();

    /**
     * 列名
     * @return
     */
    String name();

    /**
     * 日期默认格式
     * @return
     */
    String dateFormat() default "yyyy-MM-dd HH:mm:ss" ;
}
