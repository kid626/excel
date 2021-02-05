package com.bruce.excel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.bruce.excel.entity.EasyExcelData;
import com.bruce.excel.listener.EasyExcelDataListener;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
public class EasyExcelUtil {

    /**
     * 最简单的读
     * <p>1. 创建excel对应的实体对象 参照{@link EasyExcelData}
     * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link EasyExcelDataListener}
     * <p>3. 直接读即可
     *
     * @param fileName           文件名
     * @param pojoClass          对应实体类
     * @param pojoReaderListener 对应实体类监听器
     * @param headRowNumber      行头数量，决定了从哪一行开始读
     * @param <T>                实体类
     */
    public static <T> void simpleRead(String fileName, Class<T> pojoClass, ReadListener<T> pojoReaderListener, Integer headRowNumber, Integer sheetNo) {
        if (headRowNumber == null) {
            headRowNumber = 1;
        }
        if (sheetNo == null) {
            sheetNo = 0;
        }
        EasyExcel.read(fileName, pojoClass, pojoReaderListener).sheet(sheetNo).headRowNumber(headRowNumber).doRead();
    }

    /**
     * 最简单的读
     * <p>1. 创建excel对应的实体对象 参照{@link EasyExcelData}
     * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link EasyExcelDataListener}
     * <p>3. 直接读即可
     *
     * @param file               MultipartFile
     * @param pojoClass          对应实体类
     * @param pojoReaderListener 对应实体类监听器
     * @param headRowNumber      行头数量，决定了从哪一行开始读
     * @param <T>                实体类
     * @throws IOException IOException
     */
    public static <T> void simpleRead(MultipartFile file, Class<T> pojoClass, ReadListener<T> pojoReaderListener, Integer headRowNumber, Integer sheetNo) throws IOException {
        if (headRowNumber == null) {
            headRowNumber = 1;
        }
        if (sheetNo == null) {
            sheetNo = 0;
        }
        EasyExcel.read(file.getInputStream(), pojoClass, pojoReaderListener).sheet(sheetNo).headRowNumber(headRowNumber).doRead();
    }

    /**
     * 最简单的写,若要自定义格式，合并单元格请参照 {@link "https://alibaba-easyexcel.github.io/quickstart/write.html#%E8%87%AA%E5%AE%9A%E4%B9%89%E6%A0%B7%E5%BC%8F"}
     * <p>1. 创建excel对应的实体对象 参照{@link EasyExcelData}
     * <p>2. 直接写即可
     *
     * @param fileName         文件
     * @param templateFileName 模板文件
     * @param pojoClass        对应实体类
     * @param data             数据
     * @param needHead         是否需要导出表头
     * @param sheetNo          表单 no
     * @param sheetName        表单名称
     * @param <T>              实体类
     */
    public static <T> void simpleWrite(String fileName, String templateFileName, Class<T> pojoClass, List<T> data, Integer sheetNo, String sheetName, Boolean needHead) {
        if (sheetNo == null) {
            sheetNo = 0;
        }
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = "Sheet" + sheetNo;
        }
        ExcelWriterBuilder write = EasyExcel.write(fileName, pojoClass);
        if (StringUtils.isNotEmpty(templateFileName)) {
            write.withTemplate(templateFileName);
        }
        if (needHead != null) {
            write.needHead(needHead);
        }
        ExcelWriterSheetBuilder builder = write.sheet(sheetNo, sheetName);
        builder.doWrite(data);
    }

    /**
     * 最简单的写,若要自定义格式，合并单元格请参照 {@link "https://alibaba-easyexcel.github.io/quickstart/write.html#%E8%87%AA%E5%AE%9A%E4%B9%89%E6%A0%B7%E5%BC%8F"}
     * <p>1. 创建excel对应的实体对象 参照{@link EasyExcelData}
     * <p>2. 直接写即可
     *
     * @param response         HttpServletResponse
     * @param templateFileName 模板文件
     * @param pojoClass        对应实体类
     * @param data             数据
     * @param needHead         是否需要导出表头
     * @param sheetNo          表单 no
     * @param sheetName        表单名称
     * @param <T>              实体类
     */
    public static <T> void simpleWrite(HttpServletResponse response, String templateFileName, Class<T> pojoClass, List<T> data, Integer sheetNo, String sheetName, Boolean needHead) throws IOException {
        if (sheetNo == null) {
            sheetNo = 0;
        }
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = "Sheet" + sheetNo;
        }
        ExcelWriterBuilder write = EasyExcel.write(response.getOutputStream(), pojoClass);
        if (StringUtils.isNotEmpty(templateFileName)) {
            write.withTemplate(templateFileName);
        }
        if (needHead != null) {
            write.needHead(needHead);
        }
        ExcelWriterSheetBuilder builder = write.sheet(sheetNo, sheetName);
        builder.doWrite(data);
    }

    /**
     * 简单填充，适合 header + 列表 模式
     *
     * @param fileName         生成的文件名
     * @param templateFileName 模板文件名
     * @param map              用于填充 header 里的参数 模板中以 {name} 来替代
     * @param fillData         用于填充 列表，模板中以 {.name} 来替代
     * @param direction        方向 HORIZONTAL | VERTICAL {@link com.alibaba.excel.enums.WriteDirectionEnum},默认为 VERTICAL
     * @param <T>              列表实体类
     */
    public static <T> void simpleFill(String fileName, String templateFileName, Map<String, Object> map, List<T> fillData, WriteDirectionEnum direction) {
        if (direction == null) {
            direction = WriteDirectionEnum.VERTICAL;
        }
        FillConfig fillConfig = FillConfig.builder().direction(direction).build();
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollectionUtils.isNotEmpty(fillData)) {
            excelWriter.fill(fillData, fillConfig, writeSheet);
        }
        if (ObjectUtils.isNotEmpty(map)) {
            excelWriter.fill(map, fillConfig, writeSheet);
        }
        excelWriter.finish();
    }

    /**
     * 简单填充，适合 header + 列表 模式
     *
     * @param response         HttpServletResponse
     * @param templateFileName 模板文件名
     * @param map              用于填充 header 里的参数 模板中以 {name} 来替代
     * @param fillData         用于填充 列表，模板中以 {.name} 来替代
     * @param direction        方向 HORIZONTAL | VERTICAL {@link com.alibaba.excel.enums.WriteDirectionEnum},默认为 VERTICAL
     * @param <T>              列表实体类
     */
    public static <T> void simpleFill(HttpServletResponse response, String templateFileName, Map<String, Object> map, List<T> fillData, WriteDirectionEnum direction) throws IOException {
        if (direction == null) {
            direction = WriteDirectionEnum.VERTICAL;
        }
        FillConfig fillConfig = FillConfig.builder().direction(direction).build();
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollectionUtils.isNotEmpty(fillData)) {
            excelWriter.fill(fillData, fillConfig, writeSheet);
        }
        if (ObjectUtils.isNotEmpty(map)) {
            excelWriter.fill(map, fillConfig, writeSheet);
        }
        excelWriter.finish();
    }

    /**
     * 复杂填充，适合 header + 列表 + tail 模式
     *
     * @param fileName         生成的文件名
     * @param templateFileName 模板文件名
     * @param map              用于填充 header 里的参数 模板中以 {name} 来替代
     * @param fillData         用于填充 列表，模板中以 {.name} 来替代
     * @param tailData         用于写入 tail 里
     * @param <T>              列表实体类
     * @param <V>              tail 实体类
     */
    public static <T, V> void complexFill(String fileName, String templateFileName, Map<String, Object> map, List<T> fillData, List<V> tailData) {
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollectionUtils.isNotEmpty(fillData)) {
            excelWriter.fill(fillData, writeSheet);
        }
        if (ObjectUtils.isNotEmpty(map)) {
            excelWriter.fill(map, writeSheet);
        }
        excelWriter.write(tailData, writeSheet);
        excelWriter.finish();
    }

    /**
     * 复杂填充，适合 header + 列表 + tail 模式
     *
     * @param response         HttpServletResponse
     * @param templateFileName 模板文件名
     * @param map              用于填充 header 里的参数 模板中以 {name} 来替代
     * @param fillData         用于填充 列表，模板中以 {.name} 来替代
     * @param tailData         用于写入 tail 里
     * @param <T>              列表实体类
     * @param <V>              tail 实体类
     */
    public static <T, V> void complexFill(HttpServletResponse response, String templateFileName, Map<String, Object> map, List<T> fillData, List<V> tailData) throws IOException {
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollectionUtils.isNotEmpty(fillData)) {
            excelWriter.fill(fillData, writeSheet);
        }
        if (ObjectUtils.isNotEmpty(map)) {
            excelWriter.fill(map, writeSheet);
        }
        excelWriter.write(tailData, writeSheet);
        excelWriter.finish();
    }

}
