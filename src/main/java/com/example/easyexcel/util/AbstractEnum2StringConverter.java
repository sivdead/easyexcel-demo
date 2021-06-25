package com.example.easyexcel.util;

import cn.hutool.core.util.TypeUtil;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.google.common.collect.BiMap;
import com.google.common.collect.HashBiMap;

/**
 * 枚举值转换器
 * @author 92339
 */
public abstract class AbstractEnum2StringConverter<T, E extends IExcelEnum<T>> implements Converter<T> {

    private final Class<T> typeClass;
    private final BiMap<T, String> enumMap = HashBiMap.create();

    @SuppressWarnings({"unchecked"})
    public AbstractEnum2StringConverter() {
        Class<IExcelEnum<T>> enumClass = (Class<IExcelEnum<T>>) TypeUtil.getTypeArgument(getClass(), 1);
        if (!enumClass.isEnum()) {
            throw new IllegalArgumentException("ParameterizedType[1] must be enum");
        }
        this.typeClass = (Class<T>) TypeUtil.getTypeArgument(getClass(), 0);
        initEnumMap(enumClass);
    }

    private void initEnumMap(Class<IExcelEnum<T>> enumClass) {
        for (IExcelEnum<T> enumConstant : enumClass.getEnumConstants()) {
            this.enumMap.put(enumConstant.getCode(), enumConstant.getStringValue());
        }
    }

    @Override
    public Class<?> supportJavaTypeKey() {
        return typeClass;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public T convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        String stringValue = cellData.getStringValue();
        if (stringValue == null) {
            return null;
        }
        T t = enumMap.inverse().get(stringValue);
        if (t == null) {
            throw new IllegalArgumentException(String.format("invalid value in cell: %s, row: %d",
                    contentProperty.getHead().getFieldName(), cellData.getRowIndex()));
        }
        return t;
    }

    @Override
    public CellData<String> convertToExcelData(T value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {

        String stringValue = enumMap.get(value);
        if (stringValue == null) {
            throw new IllegalArgumentException(String.format("invalid value in model, fieldName: %s",
                    contentProperty.getHead().getFieldName()));
        }
        return new CellData<>(stringValue);
    }
}