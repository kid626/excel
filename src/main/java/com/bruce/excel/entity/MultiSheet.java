package com.bruce.excel.entity;

import com.alibaba.excel.read.listener.ReadListener;
import lombok.Data;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Data
public class MultiSheet {

    /**
     * sheet 编号
     */
    private Integer sheetNo;

    /**
     * 对应实体类
     */
    private Class pojoClass;

    /**
     * 实体类监听器
     */
    private ReadListener pojoReaderListener;

    /**
     * 表行头数量
     */
    private Integer headRowNumber;

}
