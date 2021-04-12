package com.zhuang.excel.easyexcel;

import com.zhuang.excel.demo.UserModel;
import com.zhuang.excel.easypoi.EasyPoiUtils;
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
    public void readToList() {
        List<UserModel> userList = EasyExcelUtils.readToList(getClass().getResource("/excel/data-test-01.xlsx").getPath(), UserModel.class);
        userList.forEach(System.out::println);
    }
}
