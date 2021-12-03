package com.bruce.excel.service.impl;

import com.bruce.excel.entity.Test;
import com.bruce.excel.listener.TestListener;
import com.bruce.excel.service.TestService;
import com.bruce.excel.util.EasyExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
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


}
