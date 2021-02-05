package com.bruce.excel.service.impl;

import com.bruce.excel.entity.Test;
import com.bruce.excel.listener.TestListener;
import com.bruce.excel.service.TestService;
import com.bruce.excel.util.EasyExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
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


}
