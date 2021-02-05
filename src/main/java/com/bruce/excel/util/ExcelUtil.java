package com.bruce.excel.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Slf4j
public class ExcelUtil {


    private final static String XLS = "xls";

    private final static String XLSX = "xlsx";

    /**
     * 导出 excel 一列一列插入
     * 适用于第一行为标题栏
     *
     * @param fileName  文件名
     * @param titleRows 标题行，中间用 "," 隔开
     * @param data      将一个表格放在一个 List<List<String>> 中
     * @param response  HttpServletResponse
     */
    public static void exportExcel(String fileName, String titleRows, List<List<String>> data, HttpServletResponse response) {
        XSSFWorkbook workbook = generateWorkbook(fileName, titleRows, data);
        try {
            ExcelUtil.buildExcelFile("C:\\Users\\fzh\\Desktop\\test3.xlsx", workbook);
            // ExcelUtil.buildExcelDocument(fileName, workbook, response);
        } catch (Exception e) {
            log.info("导出失败", e);
        }
    }

    /**
     * 导出 excel 一列一列插入
     * 适用于第一行为标题栏
     *
     * @param fileName  文件名
     * @param titleRows 标题行，中间用 "," 隔开
     * @param data      将一个表格放在一个 List<List<String>> 中
     */
    public static XSSFWorkbook generateWorkbook(String fileName, String titleRows, List<List<String>> data) {

        String[] titles = titleRows.split(",");
        // 总列数
        int columns = titles.length;
        // 总行数
        int rows = data.size();
        //存储最大列宽
        Map<Integer, Integer> maxWidth = new HashMap<>(columns);

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(fileName);

        // 初始化标题行
        XSSFRow titleRow = sheet.createRow(0);
        XSSFCellStyle cellStyle = generateCellStyle(workbook);

        for (int i = 0; i < columns; i++) {
            titleRow.createCell(i).setCellValue(titles[i]);
        }

        // 初始化标题的列宽,字体
        for (int i = 0; i < columns; i++) {
            maxWidth.put(i, titleRow.getCell(i).getStringCellValue().getBytes().length * 256 + 512);
            titleRow.getCell(i).setCellStyle(cellStyle);
        }
        for (int row = 1; row <= rows; row++) {
            XSSFRow currentRow = sheet.createRow(row);
            for (int column = 0; column < columns; column++) {
                currentRow.createCell(column).setCellValue(data.get(row - 1).get(column));
                int length = currentRow.getCell(column).getStringCellValue().getBytes().length * 256 + 512;
                //这里把宽度最大限制到15000
                if (length > 15000) {
                    length = 15000;
                }
                maxWidth.put(column, Math.max(length, maxWidth.get(column)));
                currentRow.getCell(column).setCellStyle(cellStyle);
            }
        }
        for (int i = 0; i < columns; i++) {
            sheet.setColumnWidth(i, maxWidth.get(i));
        }
        return workbook;
    }

    /**
     * 导入数据（单页）
     *
     * @param file       文件
     * @param sheetIndex 页名的索引(从0开始，-1代表全部页)
     * @return 导入某个 sheet
     * @throws IOException 异常
     */
    public static List<List<String>> importExcel(MultipartFile file, int sheetIndex) throws Exception {
        Workbook workbook;
        //返回的data
        List<List<String>> data = new ArrayList<>();
        workbook = getWorkbook(file);
        //导入某一页
        if (sheetIndex > -1) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            List<List<String>> lists = importOneSheet(sheet);
            data.addAll(lists);
        } else {
            //导入全部
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
                List<List<String>> lists = importOneSheet(sheet);
                data.addAll(lists);
            }
        }
        return data;
    }

    /**
     * 导入数据（所有页）
     *
     * @param file 文件
     * @return 导入所有 sheet
     * @throws IOException 异常
     */
    public static List<List<String>> importExcel(MultipartFile file) throws Exception {
        return importExcel(file, -1);
    }

    /**
     * 将 excelDate 转成 java 的 Date
     *
     * @param excelDate ExcelDate
     * @param format    时间显示格式，如 yyyy-MM-dd
     * @return java 的 Date
     */
    public static String getDateFormat(String excelDate, String format) {
        if (StringUtils.isEmpty(excelDate)) {
            return "";
        }
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(HSSFDateUtil.getJavaDate(Double.parseDouble(excelDate)));
    }

    /**
     * 获取百分比形式
     *
     * @param percent excel 的 百分比格式
     * @param format  百分比格式，如 #.##% 表示保留两位小数
     * @return 带上百分号
     */
    public static String getPercentFormat(String percent, String format) {
        if (StringUtils.isEmpty(percent)) {
            return "";
        }
        DecimalFormat df = new DecimalFormat(format);
        return df.format(Double.parseDouble(percent));
    }

    /**
     * 设置为日期格式
     *
     * @param cell   要设置的单元格
     * @param format 格式,如 yyyy年MM月dd日
     */
    public static void setDateFormat(XSSFCell cell, String format) {
        XSSFWorkbook workbook = cell.getSheet().getWorkbook();
        XSSFCellStyle cellStyle = generateCellStyle(workbook);
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        String cellValue = getCellValue(cell);
        if (!StringUtils.isEmpty(cellValue)) {
            Date date = HSSFDateUtil.getJavaDate(Double.parseDouble(cellValue));
            cell.setCellValue(date);
        }
        cell.getSheet().setColumnWidth(cell.getColumnIndex(), format.getBytes().length * 256 + 512);
        cellStyle.setDataFormat(dataFormat.getFormat(format));
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置格式为百分比格式
     *
     * @param cell   要设置的单元格
     * @param format 格式,如 0.00%
     */
    public static void setPercentFormat(XSSFCell cell, String format) {
        XSSFCellStyle cellStyle = generateCellStyle(cell.getSheet().getWorkbook());
        String cellValue = getCellValue(cell);
        if (!StringUtils.isEmpty(cellValue)) {
            cell.setCellValue(Double.parseDouble(cellValue));
        }
        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(format));
        cell.setCellStyle(cellStyle);
    }

    /**
     * 获取一个sheet里的数据
     * 适用于第一行为标题行的情况，若有数据为空，则自动填充，保证最后得到一个 m * n 的列表
     *
     * @param sheet 某页
     * @return 表格内容
     */
    private static List<List<String>> importOneSheet(Sheet sheet) {
        List<List<String>> data = new ArrayList<>();
        //获得当前sheet的开始行,从 0 开始
        int firstRowNum = 0;
        //获得当前sheet的结束行
        int lastRowNum = sheet.getLastRowNum();
        //获得每一行的开始列，从 0 开始
        int firstCellNum = 0;
        //获得当前行的列数，以第一行的最后一列为准
        int lastCellNum = sheet.getRow(0).getLastCellNum();
        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
            //获得当前行
            Row currentRow = sheet.getRow(rowNum);
            List<String> cells = new ArrayList<>();
            //循环当前行
            for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                Cell cell = currentRow.getCell(cellNum);
                if (cell == null) {
                    cells.add("");
                } else {
                    cells.add(getCellValue(cell));
                }
            }
            data.add(cells);
        }
        return data;
    }

    /**
     * 获取workbook
     *
     * @return HSSFWorkbook 或者 XSSFWorkbook
     */
    private static Workbook getWorkbook(MultipartFile file) throws Exception {
        Workbook workbook;
        //获取文件名
        String fileName = file.getOriginalFilename();
        if (fileName == null) {
            throw new Exception("文件名有误！");
        }
        //判断文件格式
        if (fileName.endsWith(XLS)) {
            workbook = new HSSFWorkbook(file.getInputStream());
        } else if (fileName.endsWith(XLSX)) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else {
            throw new Exception("文件格式有误!");
        }
        return workbook;
    }


    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return 均已 string 类型返回
     */
    private static String getCellValue(Cell cell) {
        String cellValue;
        // DecimalFormat df = new DecimalFormat("#");
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getRichStringCellValue().getString().trim();
                break;
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case FORMULA:
                cellValue = cell.getCellFormula();
                break;
            default:
                cellValue = "";
        }
        return cellValue.trim();
    }

    /**
     * 生成一个 通用 CellStyle（自动换行，水平垂直居中）
     *
     * @param workbook XSSFWorkbook
     * @return XSSFCellStyle
     */
    private static XSSFCellStyle generateCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        // 自动换行
        cellStyle.setWrapText(true);
        // 水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 生成本地 excel 文件
     *
     * @param fileName 文件名
     * @param workbook XSSFWorkbook
     * @throws Exception Exception
     */
    public static void buildExcelFile(String fileName, XSSFWorkbook workbook) throws Exception {
        FileOutputStream fos = new FileOutputStream(fileName);
        workbook.write(fos);
        fos.flush();
        fos.close();
    }

    /**
     * 浏览器下载excel
     *
     * @param fileName 文件名
     * @param workbook XSSFWorkbook
     * @param response HttpServletResponse
     * @throws Exception Exception
     */
    public static void buildExcelDocument(String fileName, XSSFWorkbook workbook, HttpServletResponse response) throws Exception {
        // 注释掉该行代码，可以解决 Could not find acceptable representation  问题
        // response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName, "utf-8"));
        OutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }
}
