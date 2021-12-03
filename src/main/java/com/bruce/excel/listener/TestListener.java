package com.bruce.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.bruce.excel.entity.Test;
import com.bruce.excel.service.TestService;
import com.bruce.excel.service.impl.TestServiceImpl;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Slf4j
public class TestListener extends AnalysisEventListener<Test> {

    private final List<Test> list;
    private final List<Test> errorList;

    public List<Test> getErrorList() {
        return errorList;
    }

    /**
     * 假设这个是一个DAO，当然有业务逻辑这个也可以是一个service。当然如果不用存储这个对象没用。
     */
    private final TestService testService;

    public TestListener() {
        // 这里是demo，所以随便new一个。实际使用如果到了spring,请使用下面的有参构造函数
        testService = new TestServiceImpl();
        list = new ArrayList<>();
        errorList = new ArrayList<>();
    }

    /**
     * 如果使用了spring,请使用这个构造方法。每次创建Listener的时候需要把spring管理的类传进来
     *
     * @param testService TestService
     */
    public TestListener(TestService testService) {
        this.testService = testService;
        list = new ArrayList<>();
        errorList = new ArrayList<>();
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param context AnalysisContext
     */
    @Override
    public void invoke(Test data, AnalysisContext context) {
        log.info("解析到一条数据:{}", data);
        if (data.getId() == null || StringUtils.isBlank(data.getName())
                || data.getAge() == null || StringUtils.isBlank(data.getName())) {
            errorList.add(data);
        } else {
            list.add(data);
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        log.info("所有数据解析完成！");
        testService.save(list);
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        log.error("发现异常:", exception);
    }

}
