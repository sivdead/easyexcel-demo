package com.example.easyexcel.util;

import java.lang.annotation.*;

/**
 * Excel枚举字段需要使用此注解声明
 *
 * @author jingwen
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface ExcelEnum {

    Class<? extends IExcelEnum<?>> value();

}
