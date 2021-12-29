package com.bruce.excel.controller;

import com.bruce.excel.entity.Test;
import com.bruce.excel.service.TestService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@RestController
@RequestMapping("/test")
@Slf4j
public class TestController {

    @Autowired
    private TestService testService;


    @PostMapping("/v1/import")
    public String batchSave(MultipartFile file, HttpServletResponse response) {
        if (file.isEmpty()) {
            log.info("文件为空");
            return "文件不能为空";
        }
        String originalFilename = file.getOriginalFilename();
        if (StringUtils.isBlank(originalFilename)) {
            log.info("文件名为空");
            return "文件名不能为空";
        }
        if (!(originalFilename.endsWith("xlsx") || originalFilename.endsWith("xls"))) {
            log.info("文件格式错误:{}", originalFilename);
            return "文件格式错误";
        }
        try {
            List<Test> list = testService.batchSave(file, response);
            return list.toString();
        } catch (Exception e) {
            return e.getMessage();
        }
    }

    @PostMapping("/v1/sql")
    public String sql(String prefix, MultipartFile file, HttpServletResponse response) {
        if (file.isEmpty()) {
            log.info("文件为空");
            return "文件不能为空";
        }
        String originalFilename = file.getOriginalFilename();
        if (StringUtils.isBlank(originalFilename)) {
            log.info("文件名为空");
            return "文件名不能为空";
        }
        if (!(originalFilename.endsWith("xlsx") || originalFilename.endsWith("xls"))) {
            log.info("文件格式错误:{}", originalFilename);
            return "文件格式错误";
        }
        try {
            return testService.generateSql(prefix,file, response);
        } catch (Exception e) {
            return e.getMessage();
        }
    }
}
