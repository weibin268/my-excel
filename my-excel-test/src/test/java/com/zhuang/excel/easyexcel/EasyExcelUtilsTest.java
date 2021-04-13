package com.zhuang.excel.easyexcel;

import com.zhuang.excel.demo.User4EasyExcel;
import lombok.Data;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class EasyExcelUtilsTest {

    @Data
    public static class User {
        private String name;
        private Integer age;
        private Date date;
    }

    @Test
    public void export() {
        List<User> userList = new ArrayList<>();
        User user = new User();
        user.setName("zwb");
        user.setAge(10);
        user.setDate(new Date());
        userList.add(user);
        user = new User();
        user.setName("zzz");
        user.setAge(20);
        user.setDate(new Date());
        userList.add(user);
        EasyExcelUtils.export(getClass().getResource("/excel/easyexcel-test-01.xlsx").getPath(), "D:\\temp\\easyexcel-test-01.xlsx",
                new FillItemBuilder()
                        .add("userList", userList)
                        .set("total", userList.size())
                        .buildList());

    }

    @Test
    public void export2() {
        List<User4EasyExcel> userList = new ArrayList<>();
        User4EasyExcel user = new User4EasyExcel();
        user.setName("zwb");
        user.setAge(10);
        user.setDate(new Date());
        userList.add(user);
        user = new User4EasyExcel();
        user.setName("zzz");
        user.setAge(20);
        user.setDate(new Date());
        userList.add(user);
        EasyExcelUtils.export("D:\\temp\\easyexcel-test-02.xlsx", userList, User4EasyExcel.class);
    }

    @Test
    public void readToList() {
        List<User4EasyExcel> userList = EasyExcelUtils.readToList(getClass().getResource("/excel/data-test-01.xlsx").getPath(), User4EasyExcel.class);
        userList.forEach(System.out::println);
    }
}
