package com.zhuang.excel.jxls;

import com.zhuang.excel.model.User;
import org.junit.Test;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

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

    /**
     * 合并单元格测试（注：模板中要合并的单元格不能本身就是一格合并单元格，可以模板本身的单元格不要合并，而是通过提高jxls合并单元格的数量的方式实现）
     */
    @Test
    public void export4MergeCells() {
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

        LinkedHashMap<String, List<User>> groupBySex = userList.stream().collect(Collectors.groupingBy(User::getSex, LinkedHashMap::new, Collectors.toList()));
        model.put("groupBySex", groupBySex);

        JxlsUtils.export(getClass().getResource("/excel/jxls-test-02.xlsx").getPath(), "D:\\temp\\jxls-test-02.xlsx", model);
    }

    /**
     * 表格动态列测试
     */
    @Test
    public void export4DynamicGrid() {
        Map<String, Object> model = new HashMap<>();

        List<List<Object>> userList = new ArrayList<>();
        List<Object> user = new ArrayList<>();
        user.add("zwb");
        user.add(18);
        user.add("男");
        userList.add(user);

        user = new ArrayList<>();
        user.add("lxc");
        user.add(17);
        user.add("女");
        userList.add(user);

        model.put("headers", Arrays.asList("姓名", "年龄", "性别"));
        model.put("data", userList);

        JxlsUtils.export(getClass().getResource("/excel/jxls-test-03.xlsx").getPath(), "D:\\temp\\jxls-test-03.xlsx", model);
    }

}
