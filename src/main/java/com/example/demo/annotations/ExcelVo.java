package com.example.demo.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 注解
 * @author  dqz
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelVo {

    /** 对应的列名称 */
    String name() default "";

    /** 列序号 */
    int sort();

    /** 字段类型对应的格式 */
    String dateFormat() default "";

    /** 是否需要校验 */
    boolean isCheck() default false;

    /** 校验字段长度 */
    int fieldLength() default 50;

    /** 校验是否可以为空 */
    boolean isEmpty() default true;
}
