package com.example.easyexcel.constant;

import com.example.easyexcel.util.AbstractEnum2StringConverter;
import com.example.easyexcel.util.IExcelEnum;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author 92339
 */
@AllArgsConstructor
@Getter
public enum HouseTypeEnum implements IExcelEnum<String> {

    RESIDENTIAL("residential", "住宅"),
    PARKING("parking", "车位"),
    SHOP("shop","商铺"),
    ;
    private final String code;

    private final String stringValue;

    /**
     * 声明class即可,不需要有具体实现
     */
    public static class HouseTypeEnum2StringConverter extends AbstractEnum2StringConverter<String,HouseTypeEnum>{}

}
