package com.example.easyexcel.util;

import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Arrays;
import java.util.Map;

/**
 * 具体
 * @author jingwen
 */
public class EnumConstraintSheetWriteHandler extends AbstractSheetWriteHandler {

    private final int dataSize;

    public EnumConstraintSheetWriteHandler(int dataSize) {
        this.dataSize = dataSize;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();
        DataValidationHelper helper = sheet.getDataValidationHelper();
        for (Map.Entry<Integer, ExcelContentProperty> entry : writeSheetHolder.getExcelWriteHeadProperty().getContentPropertyMap().entrySet()) {
            int index = entry.getKey();
            ExcelContentProperty excelContentProperty = entry.getValue();
            ExcelEnum excelEnum = excelContentProperty.getField().getAnnotation(ExcelEnum.class);
            if (excelEnum != null) {
                Class<? extends IExcelEnum<?>> enumClass = excelEnum.value();
                if (!enumClass.isEnum()) {
                    throw new IllegalArgumentException("ExcelEnum's value must be enum class");
                }
                String[] values = Arrays.stream(enumClass.getEnumConstants())
                        .map(IExcelEnum::getStringValue)
                        .toArray(String[]::new);
                DataValidationConstraint constraint = helper.createExplicitListConstraint(values);
                CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, dataSize + 1, index, index);
                DataValidation validation = helper.createValidation(constraint, cellRangeAddressList);
                //设置枚举列,提供下拉框,防止误操作
                sheet.addValidationData(validation);
            }
        }
    }
}
