package com.bruce.excel.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.Date;

/**
 * @Copyright Copyright © 2021 fanzh . All rights reserved.
 * @Desc
 * @ProjectName excel
 * @Date 2021/2/2 16:33
 * @Author Bruce
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class FileEntity implements Serializable {

    //注解，easypoi 教程里有，时间转换、数据字典转换等

    @Excel(name = "姓名")
    public String name;

    @Excel(name = "性别", replace = {"男_1", "女_2"}, orderNum = "1")
    public String gender;

    @Excel(name = "权", orderNum = "2")
    public String tags;

    @Excel(name = "日期", format = "yyyy-MM-dd", orderNum = "3")
    public Date date;

    @Excel(name = "图片", orderNum = "4", type = 2)
    public String pic;

}
