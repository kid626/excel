package com.bruce.excel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/12/29 0:14
 * @Author Bruce
 */
@Data
public class CommonData {

    @ExcelProperty(index = 0)
    private String str0;
    @ExcelProperty(index = 1)
    private String str1;
    @ExcelProperty(index = 2)
    private String str2;
    @ExcelProperty(index = 3)
    private String str3;
    @ExcelProperty(index = 4)
    private String str4;
    @ExcelProperty(index = 5)
    private String str5;
    @ExcelProperty(index = 6)
    private String str6;
    @ExcelProperty(index = 7)
    private String str7;
    @ExcelProperty(index = 8)
    private String str8;
    @ExcelProperty(index = 9)
    private String str9;

}
