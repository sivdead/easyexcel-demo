package com.example.easyexcel.util;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.google.common.collect.ImmutableMap;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 前提条件:
 * 1. class模式
 *
 * @author jingwen
 */
@Slf4j
public class FormulaCellWriteHandler extends AbstractCellWriteHandler implements SheetWriteHandler {

    private static final String ENV_ROW_NO = "rowNo";
    private static final Pattern PATTERN_ENV = Pattern.compile("\\$\\{([a-zA-Z]\\w+)}");
    private static final Pattern PATTERN_FUNC_COMPLEX = Pattern.compile("\\$\\{([a-zA-Z]\\w+)\\((((('[a-zA-Z]\\w+')|\\d+),?)*)\\)}");
    private static final Pattern PATTERN_FUNC_SIMPLE = Pattern.compile("([a-zA-Z]\\w*)");
    private static final Pattern PATTERN_ARGS = Pattern.compile("('([a-zA-Z]\\w*)')|(\\d+)");

    /**
     * easyExcel的元数据
     */
    private Map<String, ExcelContentProperty> contentPropertyMap = null;
    /**
     * 公式映射
     * 因为不会进行删除操作,且即使更新也是用同样的数据覆盖,故可以用HashMap
     */
    private final Map<String, ExcelFormula> formulaMap = new HashMap<>();

    private final Function<String[], String> getCellAddressFunc = (args) -> getCellAddress(args[0], Integer.parseInt(args[1]));

    private final Map<String, Function<String[], String>> functionMap = ImmutableMap.of(
            "getCellAddress", getCellAddressFunc
    );

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        //跳过头信息,如果不是class模式,也跳过
        if (!isHead && head != null && head.getFieldName() != null) {
            ExcelFormula excelFormula;
            if (formulaMap.containsKey(head.getFieldName())) {
                excelFormula = formulaMap.get(head.getFieldName());
            } else {
                ExcelContentProperty excelContentProperty = contentPropertyMap.get(head.getFieldName());
                excelFormula = excelContentProperty.getField().getAnnotation(ExcelFormula.class);
                //为空也put一下,这样下次containsKey会返回true
                formulaMap.put(head.getFieldName(), excelFormula);
            }
            if (excelFormula == null) {
                return;
            }
            String formulaTemplate = excelFormula.value();
            //替换环境变量
            String formula;
            if (excelFormula.type() == ExcelFormula.FormulaType.COMPLEX) {
                formula = parseComplexFunctionAndExecute(formulaTemplate, cell);
            } else {
                formula = parseSimpleFunctionAndExecute(formulaTemplate, cell);
            }
            cell.setCellFormula(formula);
        }
    }

    private String parseSimpleFunctionAndExecute(String formulaTemplate, Cell cell) {
        log.info("=============开始编译公式模板: {}=============", formulaTemplate);
        StringBuilder stringBuffer = new StringBuilder();
        Matcher matcher = PATTERN_FUNC_SIMPLE.matcher(formulaTemplate);
        int currentRow = cell.getAddress().getRow() + 1;
        while (matcher.find()) {
            String fieldName = matcher.group(1);
            String fieldCellAddress = getCellAddress(fieldName, currentRow);
            log.info("-----------fieldName:{},fieldCellAddress:{}-----------", fieldName, fieldCellAddress);
            matcher.appendReplacement(stringBuffer, fieldCellAddress);
        }
        matcher.appendTail(stringBuffer);
        String formula = stringBuffer.toString();
        log.info("=============公式模板解析完毕: {}=============", formula);
        return formula;
    }


    private String parseComplexFunctionAndExecute(String formulaTemplate, Cell cell) {

        formulaTemplate = setEnvProperties(formulaTemplate, cell);
        log.info("=============开始编译公式模板: {}=============", formulaTemplate);
        Matcher matcher = PATTERN_FUNC_COMPLEX.matcher(formulaTemplate);
        StringBuilder stringBuffer = new StringBuilder();
        while (matcher.find()) {
            String methodName = matcher.group(1);
            log.info("-----------methodName:{}-----------", methodName);
            Matcher argsMatcher = PATTERN_ARGS.matcher(matcher.group(2));
            //一般不会有这么多参数了
            List<String> args = new ArrayList<>(5);
            while (argsMatcher.find()) {
                //可能是字符串或者数字
                String strArgument = argsMatcher.group(2);
                String digitArgument = argsMatcher.group(3);
                String actualArgument = strArgument != null ? strArgument : digitArgument;
                log.info("arg[{}]: {}", args.size(), actualArgument);
                args.add(actualArgument);
            }
            log.info("-----------method解析完毕-----------");
            Function<String[], String> function = functionMap.get(methodName);
            if (function == null) {
                throw new IllegalArgumentException("找不到对应的方法:" + methodName);
            }
            String result = function.apply(args.toArray(new String[0]));
            log.info("execute method, result: {}", result);
            matcher.appendReplacement(stringBuffer, result);
        }
        matcher.appendTail(stringBuffer);
        String formula = stringBuffer.toString();
        log.info("=============公式模板解析完毕: {}=============", formula);
        return formula;
    }

    private String setEnvProperties(String formulaTemplate, Cell cell) {

        CellAddress cellAddress = cell.getAddress();
        int currentRow = cellAddress.getRow() + 1;
        Matcher matcher = PATTERN_ENV.matcher(formulaTemplate);
        StringBuilder stringBuffer = new StringBuilder();
        while (matcher.find()) {
            String envName = matcher.group(1);
            if (ENV_ROW_NO.equals(envName)) {
                matcher.appendReplacement(stringBuffer, String.valueOf(currentRow));
            } else {
                throw new IllegalArgumentException("无法识别的变量");
            }
        }
        matcher.appendTail(stringBuffer);
        return stringBuffer.toString();
    }


    public String getCellAddress(String fieldName, int rowIndex) {
        ExcelContentProperty excelContentProperty = contentPropertyMap.get(fieldName);
        if (excelContentProperty == null) {
            throw new IllegalArgumentException("无效的字段名:" + fieldName);
        }
        int columnIndex = excelContentProperty.getHead().getColumnIndex();
        String columnStr = CellReference.convertNumToColString(columnIndex);
        return columnStr + rowIndex;
    }


    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (writeSheetHolder.getClazz() == null) {
            throw new UnsupportedOperationException("只支持class模式写入");
        }
        initHeadMap(writeSheetHolder.getExcelWriteHeadProperty().getContentPropertyMap());
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    private synchronized void initHeadMap(Map<Integer, ExcelContentProperty> excelContentPropertyMap) {
        if (this.contentPropertyMap == null) {
            this.contentPropertyMap = excelContentPropertyMap.values()
                    .stream()
                    .collect(Collectors.toMap(
                            e -> e.getHead().getFieldName(),
                            e -> e
                    ));
        }
    }
}