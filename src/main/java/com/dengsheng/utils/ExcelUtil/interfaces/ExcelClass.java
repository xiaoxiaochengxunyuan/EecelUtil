package com.dengsheng.utils.ExcelUtil.interfaces;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelClass{
    String value() default "";
    String sheetName() default "sheet";
}
