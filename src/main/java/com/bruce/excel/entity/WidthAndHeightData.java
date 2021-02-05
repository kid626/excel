package com.bruce.excel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import lombok.Data;

import java.util.Date;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Data
@ContentRowHeight(20)
@HeadRowHeight(50)
@ColumnWidth(10)
public class WidthAndHeightData {

    @ExcelProperty("字符串标题")
    private String string;
    @ExcelProperty("日期标题")
    private Date date;

    /**
     * 宽度为50
     */
    @ColumnWidth(100)
    @ExcelProperty(value = "数字标题")
    private Double doubleData;
}
