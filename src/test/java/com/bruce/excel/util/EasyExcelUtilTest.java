package com.bruce.excel.util;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.bruce.excel.entity.*;
import com.bruce.excel.handler.CustomCellWriteHandler;
import com.bruce.excel.handler.CustomSheetWriteHandler;
import com.bruce.excel.listener.ConverterDataListener;
import com.bruce.excel.listener.EasyExcelDataListener;
import com.bruce.excel.listener.IndexOrNameDataListener;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.util.IOUtils;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.*;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Slf4j
public class EasyExcelUtilTest {


    @Test
    public void testEasyExcelImport() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("easyexcel/in/easy_excel_import_test.xlsx");
        File file = classPathResource.getFile();
        EasyExcelUtil.simpleRead(file.getAbsolutePath(), EasyExcelData.class, new EasyExcelDataListener(), 1, 0);
    }

    @Test
    public void testEasyExcelImportV2() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("easyexcel/in/easy_excel_import_test.xlsx");
        File file = classPathResource.getFile();
        EasyExcelUtil.simpleRead(file.getAbsolutePath(), IndexOrNameData.class, new IndexOrNameDataListener(), 1, 0);
    }

    @Test
    public void testEasyExcelImportV4() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("easyexcel/in/easy_excel_import_test.xlsx");
        File file = classPathResource.getFile();
        EasyExcelUtil.simpleRead(file.getAbsolutePath(), ConverterData.class, new ConverterDataListener(), 1, 0);
    }

    @Test
    public void testEasyExcelImportV5() throws Exception {
        ClassPathResource classPathResource = new ClassPathResource("easyexcel/in/easy_excel_import_test.xlsx");
        File file = classPathResource.getFile();
        FileInputStream input = new FileInputStream(file);
        MultipartFile multipartFile = new MockMultipartFile("file", file.getName(), "text/plain", IOUtils.toByteArray(input));
        EasyExcelUtil.simpleRead(multipartFile, EasyExcelData.class, new EasyExcelDataListener(), 1, 0);
    }

    @Test
    public void testEasyExcelExportV1() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export.xlsx";
        List<EasyExcelExportData> data = new ArrayList<>();
        EasyExcelExportData data1 = new EasyExcelExportData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelExportData data2 = new EasyExcelExportData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);
        EasyExcelUtil.simpleWrite(fileName, null, EasyExcelExportData.class, data, 2, "Sheet2", true);
    }

    @Test
    public void testEasyExcelExportV2() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_with_header.xlsx";
        List<EasyExcelExportWithHeaderData> data = new ArrayList<>();
        EasyExcelExportWithHeaderData data1 = new EasyExcelExportWithHeaderData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelExportWithHeaderData data2 = new EasyExcelExportWithHeaderData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);
        Set<Integer> exclude = new HashSet<>();
        exclude.add(5);
        EasyExcelUtil.simpleWrite(fileName, null, EasyExcelExportWithHeaderData.class, data, 0, "", true);
    }

    @Test
    public void testEasyExcelExportV3() throws IOException {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_img.xlsx";
        String imagePath = "src/main/resources/img/canteen_bg.png";
        InputStream inputStream;
        List<ImageData> data = new ArrayList<>();
        ImageData data1 = new ImageData();
        data1.setString(imagePath);
        data1.setByteArray(FileUtils.readFileToByteArray(new File(imagePath)));
        data1.setFile(new File(imagePath));
        inputStream = FileUtils.openInputStream(new File(imagePath));
        data1.setInputStream(inputStream);
        data1.setUrl(new URL("https://tse4-mm.cn.bing.net/th/id/OIP.yzaU2UJBbQYEee492BL2PAHaFx?pid=Api&rs=1"));
        data.add(data1);
        EasyExcelUtil.simpleWrite(fileName, null, ImageData.class, data, 0, "", true);
        if (inputStream != null) {
            inputStream.close();
        }
    }

    @Test
    public void testEasyExcelExportV4() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_template_out.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_export_template.xlsx";
        List<TemplateExportData> data = new ArrayList<>();
        TemplateExportData data1 = new TemplateExportData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        TemplateExportData data2 = new TemplateExportData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);
        EasyExcelUtil.simpleWrite(fileName, templateFileName, TemplateExportData.class, data, 0, "", false);
    }

    @Test
    public void testEasyExcelExportV5() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_width_height.xlsx";
        List<WidthAndHeightData> data = new ArrayList<>();
        WidthAndHeightData data1 = new WidthAndHeightData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        WidthAndHeightData data2 = new WidthAndHeightData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);
        EasyExcelUtil.simpleWrite(fileName, null, WidthAndHeightData.class, data, 0, "", true);
    }

    @Test
    public void testEasyExcelExportV6() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_style.xlsx";
        List<EasyExcelData> data = new ArrayList<>();
        EasyExcelData data1 = new EasyExcelData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelData data2 = new EasyExcelData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);


        // 表头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        // 字体设置
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 20);
        headWriteFont.setBold(true);
        headWriteCellStyle.setWriteFont(headWriteFont);


        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景设置为绿色
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        // 字体设置
        WriteFont contentWriteFont = new WriteFont();
        contentWriteFont.setFontHeightInPoints((short) 20);
        contentWriteCellStyle.setWriteFont(contentWriteFont);

        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, EasyExcelData.class)
                .registerWriteHandler(horizontalCellStyleStrategy).sheet("模板").doWrite(data);
    }

    @Test
    public void testEasyExcelExportV7() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_merge.xlsx";
        List<EasyExcelData> data = new ArrayList<>();
        EasyExcelData data1 = new EasyExcelData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelData data2 = new EasyExcelData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);


        LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(2, 0);
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, EasyExcelData.class)
                .registerWriteHandler(loopMergeStrategy).sheet("模板").doWrite(data);
    }

    @Test
    public void testEasyExcelExportV8() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_max_length.xlsx";
        List<EasyExcelData> data = new ArrayList<>();
        EasyExcelData data1 = new EasyExcelData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelData data2 = new EasyExcelData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);


        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, EasyExcelData.class)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).sheet("模板").doWrite(data);
    }

    @Test
    public void testEasyExcelExportV9() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_export_custome.xlsx";
        List<EasyExcelData> data = new ArrayList<>();
        EasyExcelData data1 = new EasyExcelData();
        data1.setDate(new Date());
        data1.setDoubleData(1d);
        data1.setString("1");
        data.add(data1);
        EasyExcelData data2 = new EasyExcelData();
        data2.setDate(new Date());
        data2.setDoubleData(2d);
        data2.setString("2");
        data.add(data2);

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, EasyExcelData.class)
                .registerWriteHandler(new CustomCellWriteHandler())
                .registerWriteHandler(new CustomSheetWriteHandler()).sheet("模板").doWrite(data);
    }

    @Test
    public void testEasyExcelFillExportV1() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_fill_export_simple.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_fill_template.xlsx";
        Map<String, Object> map = new HashMap<>();
        map.put("title1", "姓名");
        map.put("title2", "年龄");
        map.put("title3", "姓名+年龄");
        map.put("title4", "忽略");
        EasyExcelUtil.simpleFill(fileName, templateFileName, map, null, null);
    }


    @Test
    public void testEasyExcelFillExportV2() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_fill_export_table.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_fill_template.xlsx";
        List<FillData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            fillData.setName("张三" + i);
            fillData.setNumber(i);
            list.add(fillData);
        }
        EasyExcelUtil.simpleFill(fileName, templateFileName, null, list, null);
    }


    @Test
    public void testEasyExcelFillExportV3() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_fill_export_simple_table.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_fill_template.xlsx";
        List<FillData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            fillData.setName("张三" + i);
            fillData.setNumber(i);
            list.add(fillData);
        }
        Map<String, Object> map = new HashMap<>();
        map.put("title1", "姓名");
        map.put("title2", "年龄");
        map.put("title3", "姓名+年龄");
        map.put("title4", "忽略");
        EasyExcelUtil.simpleFill(fileName, templateFileName, map, list, null);
    }

    @Test
    public void testEasyExcelFillExportV4() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_fill_export_complex.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_fill_template.xlsx";
        List<FillData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            fillData.setName("张三" + i);
            fillData.setNumber(i);
            list.add(fillData);
        }
        Map<String, Object> map = new HashMap<>();
        map.put("title1", "姓名");
        map.put("title2", "年龄");
        map.put("title3", "姓名+年龄");
        map.put("title4", "忽略");
        List<List<String>> tailData = new ArrayList<>();
        List<String> totalList = new ArrayList<>();
        tailData.add(totalList);
        totalList.add(null);
        totalList.add(null);
        totalList.add(null);
        totalList.add("统计:1000");
        EasyExcelUtil.complexFill(fileName, templateFileName, map, list, tailData);
    }

    @Test
    public void testEasyExcelFillExportV5() {
        String fileName = "src/main/resources/easyexcel/out/easy_excel_fill_export_horizontal.xlsx";
        String templateFileName = "src/main/resources/easyexcel/in/easy_excel_fill_template_ horizontal.xlsx";
        List<FillData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            fillData.setName("张三" + i);
            fillData.setNumber(i);
            list.add(fillData);
        }
        Map<String, Object> map = new HashMap<>();
        map.put("date", new Date());
        EasyExcelUtil.simpleFill(fileName, templateFileName, map, list, WriteDirectionEnum.HORIZONTAL);
    }

}