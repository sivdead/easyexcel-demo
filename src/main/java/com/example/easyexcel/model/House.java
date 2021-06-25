package com.example.easyexcel.model;

import com.example.easyexcel.constant.HouseTypeEnum;
import lombok.*;

import java.math.BigDecimal;

/**
 * @author 92339
 */
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class House {

    /**
     * 房间id
     */
    private Integer id;

    /**
     * 房间名称
     */
    private String houseName;

    /**
     * 面积
     */
    private BigDecimal area;

    /**
     * 单价
     */
    private BigDecimal unitPrice;

    /**
     * 总价
     */
    private BigDecimal totalPrice;

    /**
     * 业态
     * @see com.example.easyexcel.constant.HouseTypeEnum
     */
    private String type;

}
