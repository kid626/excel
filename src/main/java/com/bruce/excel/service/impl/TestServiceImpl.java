package com.bruce.excel.service.impl;

import com.bruce.excel.entity.CommonData;
import com.bruce.excel.entity.Test;
import com.bruce.excel.listener.CommonListener;
import com.bruce.excel.listener.TestListener;
import com.bruce.excel.service.TestService;
import com.bruce.excel.util.ChineseToFirstLetterUtil;
import com.bruce.excel.util.EasyExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Service
@Slf4j
public class TestServiceImpl implements TestService {

    @Override
    public List<Test> batchSave(MultipartFile file, HttpServletResponse response) throws Exception {
        // 用来接收表内所有数据
        TestListener listener = new TestListener(this);
        EasyExcelUtil.simpleRead(file, Test.class, listener, 1, 0);
        List<Test> errorList = listener.getErrorList();
        log.info("errorList.size:{}", errorList.size());
        return errorList;
    }

    @Override
    public void save(List<Test> data) {
        List<String> names = data.stream().map(Test::getName).collect(Collectors.toList());
        if (names.size() != data.size()) {
            log.info("姓名重复!");
            throw new RuntimeException("姓名重复!");
        }
        log.info("data.size:{}", data.size());
    }

    public static String DDL = "CREATE TABLE `{0}`  (\n" +
            "  `id` int(0) NOT NULL AUTO_INCREMENT COMMENT ''主键'',\n" +
            "{1}" +
            "  PRIMARY KEY (`id`)\n" +
            ") COMMENT = ''{2}'';";

    public static String COLUMN = "  `{0}` varchar(255) NULL COMMENT ''{1}'',\n";

    @Override
    public String generateSql(String prefix, MultipartFile file, HttpServletResponse response) throws Exception {
        String originalFilename = file.getOriginalFilename().replaceAll(".xlsx", "").replaceAll(".xls", "");
        String tableName = ChineseToFirstLetterUtil.chineseToFirstLetter(originalFilename);
        tableName = ChineseToFirstLetterUtil.chineseToFirstLetter(prefix) + "_" + tableName;
        List<CommonData> list = new ArrayList<>();
        EasyExcelUtil.simpleRead(file, CommonData.class, new CommonListener(list), 0, 0);
        // 第一行为标题行
        CommonData title = list.get(0);
        // 获取所有字段
        Field[] declaredFields = CommonData.class.getDeclaredFields();
        List<String> columns = new ArrayList<>();
        for (Field field : declaredFields) {
            // 设置为可访问
            field.setAccessible(true);
            // 获取对应的值
            Object value = field.get(title);
            if (value != null) {
                String valueStr = value.toString();
                columns.add(valueStr);
            }
        }
        StringBuilder sb = new StringBuilder();
        for (String column : columns) {
            sb.append(MessageFormat.format(COLUMN,
                    ChineseToFirstLetterUtil.chineseToFirstLetter(column),
                    ChineseToFirstLetterUtil.chineseLetter(column)));
        }
        StringBuilder result = new StringBuilder();
        result.append("# 初始化脚本 DDL:");
        String ddl = MessageFormat.format(DDL, tableName, sb.toString(), prefix + "-" + originalFilename);
        result.append(ddl);

        return result.toString();
    }


}
