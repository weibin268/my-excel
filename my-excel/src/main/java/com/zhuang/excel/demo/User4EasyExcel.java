package com.zhuang.excel.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

import java.util.Date;

@Data
public class User4EasyExcel {

    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private Integer age;
    @ColumnWidth(18)
    @ExcelProperty("生日")
    private Date date;

}
