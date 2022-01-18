package com.zhuang.excel.hutool;

import com.zhuang.excel.easyexcel.EasyExcelUtilsTest;
import com.zhuang.excel.model.ExportOption;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.junit.Assert.*;

public class HutoolExcelUtilsTest {

    @Test
    public void export() {
        List<EasyExcelUtilsTest.User> userList = new ArrayList<>();
        EasyExcelUtilsTest.User user = new EasyExcelUtilsTest.User();
        user.setName("zwb");
        user.setAge(10);
        user.setDate(new Date());
        userList.add(user);
        user = new EasyExcelUtilsTest.User();
        user.setName("zzz");
        user.setAge(20);
        user.setDate(new Date());
        userList.add(user);
        ExportOption exportOption = new ExportOption();
        exportOption.addColumn("name", "姓名").addColumn("age", "年龄", 50);
        HutoolExcelUtils.export("/Users/zhuang/Documents/temp/test.xlsx", userList, exportOption);
    }
}