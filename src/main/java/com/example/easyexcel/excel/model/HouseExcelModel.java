package com.example.easyexcel.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.example.easyexcel.constant.HouseTypeEnum;
import com.example.easyexcel.util.ExcelEnum;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.math.BigDecimal;

/**
 * @author 92339
 */
@Getter
@Setter
@HeadStyle(horizontalAlignment = HorizontalAlignment.CENTER, locked = true)
@ContentStyle(horizontalAlignment = HorizontalAlignment.CENTER,locked = true)
public class HouseExcelModel {

    /**
     * 房间id
     */
    @ExcelProperty("id")
    @ColumnWidth(0)
    private Integer id;

    /**
     * 房间名称
     */
    @ExcelProperty("房间名称")
    @ColumnWidth(30)
    private String houseName;

    /**
     * 面积
     */
    @ExcelProperty("面积")
    private BigDecimal area;

    /**
     * 单价
     */
    @ExcelProperty("单价(元/平方米)")
    @ContentStyle(horizontalAlignment = HorizontalAlignment.CENTER,locked = false)
    private BigDecimal unitPrice;

    /**
     * 总价
     */
    @ExcelProperty("总价(元)")
    @ContentStyle(horizontalAlignment = HorizontalAlignment.CENTER,locked = false)
    private BigDecimal totalPrice;

    /**
     * 业态
     *
     * @see com.example.easyexcel.constant.HouseTypeEnum
     */
    @ExcelProperty(value = "业态",converter = HouseTypeEnum.HouseTypeEnum2StringConverter.class)
    @ContentStyle(horizontalAlignment = HorizontalAlignment.CENTER,locked = false)
    @ColumnWidth(15)
    @ExcelEnum(HouseTypeEnum.class)
    private String type;

}
