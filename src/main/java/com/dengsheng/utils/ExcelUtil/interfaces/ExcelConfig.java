package com.dengsheng.utils.ExcelUtil.interfaces;

import com.dengsheng.utils.ExcelUtil.enums.EmptyEums;

import java.lang.annotation.*;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelConfig{
    String value() default "";
    short index() default -1;
    Class enumClass() default EmptyEums.class;
    String enumFieldName() default "";
    boolean igNore() default false;
}
