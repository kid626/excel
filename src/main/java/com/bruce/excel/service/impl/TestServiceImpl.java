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

    private List<Test> list = new ArrayList<>();

    @Override
    public String batchSave(MultipartFile file, HttpServletResponse response) throws Exception {
        // 用来接收表内所有数据
        List<Test> result = new ArrayList<>();
        EasyExcelUtil.simpleRead(file, Test.class, new TestListener(result), 9, 0);
        log.info("result.size:{}", result.size());
        save(result);
        return "";
    }

    @Override
    public void save(List<Test> data) throws Exception {
        List<String> names = list.stream().map(test -> {
            return test.getName();
        }).collect(Collectors.toList());
        for (Test test : data) {
            if (names.contains(test.getName())) {
                log.info("姓名已存在!");
                throw new Exception("姓名已存在!");
            }
            list.add(test);
            names.add(test.getName());
        }
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
        String createDDL = MessageFormat.format(DDL, tableName, sb.toString(), prefix + "-" + originalFilename);


        return createDDL;
    }


}
