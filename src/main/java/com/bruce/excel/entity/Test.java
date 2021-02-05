package com.bruce.excel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.bruce.excel.convert.CustomStringStringConverter;
import lombok.Data;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Data
public class Test {

    @ExcelProperty(index = 0)
    private Long id;

    @ExcelProperty(index = 1, converter = CustomStringStringConverter.class)
    private String name;

    @ExcelProperty(index = 2)
    private Integer age;

    @ExcelProperty(index = 3)
    private String date;

}
