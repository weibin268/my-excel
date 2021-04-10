package com.zhuang.excel.easyexcel;

import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItemBuilder;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import static org.junit.Assert.*;

public class EasyExcelUtilsTest {


    public static class User {
        private String name;
        private Integer age;
        private Date date;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }

        public Date getDate() {
            return date;
        }

        public void setDate(Date date) {
            this.date = date;
        }
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
}
