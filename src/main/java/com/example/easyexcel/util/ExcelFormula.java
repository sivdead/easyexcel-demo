package com.example.easyexcel.util;

import java.lang.annotation.*;

/**
 * @author jingwen
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface ExcelFormula {

    /**
     * 对应公式模板
     * 可以使用预制参数如${RowNo}表示当前行号,预制函数如${GetCellAddress("id",${RowNo})}获取单元格地址,可用的参数有
     * RowNo 当前行号
     */
    String value();

    FormulaType type() default FormulaType.SIMPLE;


    enum FormulaType {
        /**
         * 简单模式,不支持调用excel函数,只支持加减乘除括号
         * 但可以直接使用变量名来指向该行的某个字段
         */
        SIMPLE,
        /**
         * 复杂模式,支持excel函数调用,支持引入自定义的环境变量
         */
        COMPLEX,
    }

}
