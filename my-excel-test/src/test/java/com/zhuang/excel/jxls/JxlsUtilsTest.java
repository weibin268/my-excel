package com.zhuang.excel.jxls;

import com.zhuang.excel.model.User;
import org.junit.Test;

import java.util.*;

public class JxlsUtilsTest {

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
        model.put("list", userList);
        JxlsUtils.export(getClass().getResource("/excel/jxls-test-01.xlsx").getPath(), "D:\\temp\\jxls-test-01.xlsx", model);
    }

    @Test
    public void export2() {
        Map<String, Object> model = new HashMap<>();
        List<User> userList = new ArrayList<>();
        User user = new User();
        user.setName("zwb");
        user.setAge(10);
        user.setDate(new Date());
        user.setSex("男");
        userList.add(user);

        user = new User();
        user.setName("aaa");
        user.setAge(20);
        user.setDate(new Date());
        user.setSex("男");
        userList.add(user);

        user = new User();
        user.setName("bbb");
        user.setAge(20);
        user.setDate(new Date());
        user.setSex("女");
        userList.add(user);

        user = new User();
        user.setName("ccc");
        user.setAge(20);
        user.setDate(new Date());
        user.setSex("女");
        userList.add(user);
        model.put("list", userList);
        JxlsUtils.export(getClass().getResource("/excel/jxls-test-02.xlsx").getPath(), "D:\\temp\\jxls-test-02.xlsx", model);
    }
}
