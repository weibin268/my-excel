package com.zhuang.excel.easypoi;

import lombok.Data;
import org.junit.Test;

import java.util.*;

public class EasyPoiUtilsTest {

    @Data
    public static class User {
        private String name;
        private Integer age;
        private Date date;
    }

    @Test
    public void export() {
        Map<String, Object> model = new HashMap<>();
        List<EasyPoiUtilsTest.User> userList = new ArrayList<>();
        EasyPoiUtilsTest.User user = new EasyPoiUtilsTest.User();
        user.setName("zwb");
        user.setAge(10);
        user.setDate(new Date());
        userList.add(user);
        user = new EasyPoiUtilsTest.User();
        user.setName("zzz");
        user.setAge(20);
        user.setDate(new Date());
        userList.add(user);
        model.put("userList", userList);
        model.put("total", 100);
        EasyPoiUtils.export(getClass().getResource("/excel/easypoi-test-01.xlsx").getPath(), "D:\\temp\\easypoi-test-01.xlsx", model);
    }


}
