package com.bruce.excel.util;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Slf4j
public class ExcelUtilTest {

    @Test
    public void importExcel() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("excel/in/import_test.xlsx");
        File file = classPathResource.getFile();
        FileInputStream input = new FileInputStream(file);
        MultipartFile multipartFile = new MockMultipartFile("file", file.getName(), "text/plain", IOUtils.toByteArray(input));
        List<List<String>> lists = ExcelUtil.importExcel(multipartFile);
        lists.remove(0);
        for (List<String> list : lists) {
            System.out.print("测试数字:" + list.get(0));
            System.out.print("测试日期:" + ExcelUtil.getDateFormat(list.get(1), "yyyy-MM-dd"));
            System.out.print("测试百分比:" + ExcelUtil.getPercentFormat(list.get(2), "#.##%"));
            System.out.print("测试字符串:" + list.get(3));
            System.out.print("测试枚举:" + list.get(4));
            System.out.println();
        }
    }

    @Test
    public void exportExcel() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("excel/in/import_test.xlsx");
        File file = classPathResource.getFile();
        FileInputStream input = new FileInputStream(file);
        MultipartFile multipartFile = new MockMultipartFile("file", file.getName(), "text/plain", IOUtils.toByteArray(input));
        List<List<String>> data = ExcelUtil.importExcel(multipartFile);
        data.remove(0);
        XSSFWorkbook workbook = ExcelUtil.generateWorkbook("test", "测试数字,测试日期,测试百分比,测试字符串,测试枚举", data);
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (int row = 1; row <= sheet.getLastRowNum(); row++) {
            XSSFCell cell = sheet.getRow(row).getCell(1);
            ExcelUtil.setDateFormat(cell, "yyyy年MM月dd日 HH时mm分ss秒");
            XSSFCell cell2 = sheet.getRow(row).getCell(2);
            ExcelUtil.setPercentFormat(cell2, "0.00%");
        }
        try {
            ExcelUtil.buildExcelFile("src/main/resources/excel/out/export_test.xlsx", workbook);
        } catch (Exception e) {
            log.info("导出失败", e);
        }
    }

}