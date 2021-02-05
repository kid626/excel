package com.bruce.excel.util;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
public class EasyPOIUtil {

    /**
     * 导出 excel
     *
     * @param list      对应某实体类的 list
     * @param title     标题（与 header 区分）
     * @param sheetName sheet 的名称
     * @param pojoClass 实体类
     * @param fileName  文件名
     * @param response  HttpServletResponse
     * @throws Exception exception
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, HttpServletResponse response) throws Exception {
        ExportParams exportParams = new ExportParams(title, sheetName, ExcelType.XSSF);
        exportParams.setHeight((short) 10);
        exportParams.setTitleHeight((short) 10);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        // downLoadExcel(fileName, response, workbook);
        buildExcelFile(fileName, workbook);
    }

    /**
     * 导出多个 sheet
     *
     * @param list     对应多个 sheet
     *                 每个 map 对应一个 sheet，必须的有 3 个键
     *                 title 需要 ExportParams
     *                 entity 实体类
     *                 data   List<?>
     * @param fileName 文件名
     * @param response HttpServletResponse
     * @throws Exception
     */
    public static void exportExcel(List<Map<String, Object>> list, String fileName, HttpServletResponse response) throws Exception {
        Workbook workbook = ExcelExportUtil.exportExcel(list, ExcelType.XSSF);
        // downLoadExcel(fileName, response, workbook);
        buildExcelFile(fileName, workbook);
    }

    /**
     * 下载
     *
     * @param fileName 文件名
     * @param response HttpServletResponse
     * @param workbook Workbook
     * @throws Exception exception
     */
    private static void downLoadExcel(String fileName, HttpServletResponse response, Workbook workbook) throws Exception {
        try {
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }
    }

    /**
     * 下载到本地
     *
     * @param fileName 文件名
     * @param workbook Workbook
     * @throws Exception exception
     */
    private static void buildExcelFile(String fileName, Workbook workbook) throws Exception {
        FileOutputStream fos = new FileOutputStream(fileName);
        workbook.write(fos);
        fos.flush();
        fos.close();
    }


    /**
     * 导入 excel
     *
     * @param filePath   文件路径
     * @param titleRows  表格标题行数,默认0
     * @param headerRows 表头行数,默认1
     * @param pojoClass  导入接收的实体类
     * @return list
     * @throws Exception Exception
     */
    public static <T> List<T> importExcel(String filePath, Integer titleRows, Integer headerRows, Class<T> pojoClass) throws Exception {
        if (StringUtils.isBlank(filePath)) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list;
        try {
            list = ExcelImportUtil.importExcel(new File(filePath), pojoClass, params);
        } catch (NoSuchElementException e) {
            throw new Exception("模板不能为空");
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception(e.getMessage());
        }
        return list;
    }

    /**
     * 导入 excel
     *
     * @param file       MultipartFile
     * @param titleRows  表格标题行数,默认0
     * @param headerRows 表头行数,默认1
     * @param pojoClass  导入接收的实体类
     * @return list
     * @throws Exception Exception
     */
    public static <T> List<T> importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class<T> pojoClass) throws Exception {
        if (file == null) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list;
        try {
            list = ExcelImportUtil.importExcel(file.getInputStream(), pojoClass, params);
        } catch (NoSuchElementException e) {
            throw new Exception("excel文件不能为空");
        } catch (Exception e) {
            throw new Exception(e.getMessage());
        }
        return list;
    }

}
