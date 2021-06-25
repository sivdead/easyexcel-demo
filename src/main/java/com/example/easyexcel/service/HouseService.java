package com.example.easyexcel.service;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.lang.UUID;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.example.easyexcel.excel.model.HouseExcelModel;
import com.example.easyexcel.model.House;
import com.example.easyexcel.util.EnumConstraintSheetWriteHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.stereotype.Service;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import static com.example.easyexcel.constant.HouseTypeEnum.RESIDENTIAL;
import static com.example.easyexcel.constant.HouseTypeEnum.SHOP;

/**
 * @author 92339
 */
@Service
@Slf4j
public class HouseService {

    private final FileService fileService;

    private final static List<House> HOUSE_LIST = new ArrayList<>(50);

    static {
        HOUSE_LIST.add(new House(1, "万科城市之光-一期-一栋-101", new BigDecimal("50"), null, null, SHOP.getCode()));
        HOUSE_LIST.add(new House(2, "万科城市之光-一期-一栋-102", new BigDecimal("50"), null, null, SHOP.getCode()));
        HOUSE_LIST.add(new House(3, "万科城市之光-一期-一栋-103", new BigDecimal("50"), null, null, SHOP.getCode()));
        HOUSE_LIST.add(new House(4, "万科城市之光-一期-一栋-201", new BigDecimal("24.5"), null, null, RESIDENTIAL.getCode()));
        HOUSE_LIST.add(new House(5, "万科城市之光-一期-一栋-202", new BigDecimal("35"), null, null, RESIDENTIAL.getCode()));
        HOUSE_LIST.add(new House(6, "万科城市之光-一期-一栋-203", new BigDecimal("31"), null, null, RESIDENTIAL.getCode()));
    }

    public HouseService(FileService fileService) {
        this.fileService = fileService;
    }


    public List<House> queryHouseList() {
        return HOUSE_LIST;
    }

    @SneakyThrows
    public String export2Excel() {
        String filename = "房间列表";
        String extName = ".xlsx";
        File tempFile = File.createTempFile(filename, extName);
        log.info("temp file path: {}", tempFile.getAbsolutePath());

        List<HouseExcelModel> houseExcelModelList = queryHouseList()
                .stream()
                .map(house -> {
                    HouseExcelModel houseExcelModel = new HouseExcelModel();
                    BeanUtil.copyProperties(house, houseExcelModel);
                    return houseExcelModel;
                })
                .collect(Collectors.toList());

        //写入excel
        EasyExcel.write(tempFile)
                .head(HouseExcelModel.class)
                .sheet("房间列表")
                .registerWriteHandler(
                        new AbstractSheetWriteHandler() {
                            @Override
                            public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
                                Sheet sheet = writeSheetHolder.getSheet();
                                sheet.protectSheet(UUID.fastUUID().toString(true));
                                if (sheet instanceof XSSFSheet) {
                                    ((XSSFSheet) sheet).enableLocking();
                                } else if (sheet instanceof SXSSFSheet) {
                                    ((SXSSFSheet) sheet).enableLocking();
                                }
                            }
                        }
                )
                .registerWriteHandler(new EnumConstraintSheetWriteHandler(houseExcelModelList.size()))
                .doWrite(houseExcelModelList);

        //上传到oss,返回url给前端
        return fileService.upload(tempFile, filename + extName);
    }
}
