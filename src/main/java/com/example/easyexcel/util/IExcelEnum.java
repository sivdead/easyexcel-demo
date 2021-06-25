package com.example.easyexcel.util;

/**
 * 需要转换的枚举类需要继承此接口
 * @author 92339
 */
public interface IExcelEnum<T> {

    T getCode();

    String getStringValue();

}
