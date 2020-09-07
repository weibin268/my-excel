package com.zhuang.excel.jxls;

import org.junit.Test;

import java.util.*;

import static org.junit.Assert.*;

public class JxlsUtilsTest {

    public static class User{
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
        Map<String, Object> model = new HashMap<>();
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
        model.put("list",userList);
        JxlsUtils.export(getClass().getResource("/excel/jxls-test-01.xlsx").getPath(),"D:\\temp\\jxls-test-01.xlsx",model);
    }
}