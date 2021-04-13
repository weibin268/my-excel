package com.zhuang.excel.demo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.util.Date;

@Data
public class User4EasyPoi {

    @Excel(name = "姓名")
    private String name;
    @Excel(name = "年龄")
    private Integer age;
    @Excel(name = "生日", width = 18, format = "yyyy-MM-dd HH:mm:ss")
    private Date date;

}
