package com.bruce.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.bruce.excel.entity.CommonData;
import lombok.extern.slf4j.Slf4j;

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
public class CommonListener extends AnalysisEventListener<CommonData> {

    private List<CommonData> list = new ArrayList<>();


    public CommonListener() {
    }

    /**
     * 需要接收 list 的内容
     *
     * @param list
     */
    public CommonListener(List<CommonData> list) {
        this.list = list;
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param context
     */
    @Override
    public void invoke(CommonData data, AnalysisContext context) {
        log.info("解析到一条数据:{}", data);
        list.add(data);
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
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        log.error("发现异常:");
        throw exception;
    }

}
