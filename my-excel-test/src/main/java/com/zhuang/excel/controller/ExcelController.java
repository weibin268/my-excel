package com.zhuang.excel.controller;

import com.zhuang.excel.demo.User4EasyExcel;
import com.zhuang.excel.demo.User4EasyPoi;
import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItemBuilder;
import com.zhuang.excel.model.User;
import com.zhuang.excel.util.ExcelUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.*;

@Controller
@RequestMapping(value = "/excel")
public class ExcelController {

    @ResponseBody
    @GetMapping(value = "test")
    public String test() {
        return "{test:123}";
    }

    @RequestMapping(value = "export4EasyExcel")
    public void export4EasyExcel() {
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
        ExcelUtils.export4EasyExcel("/excel/easyexcel-test-01.xlsx", "easyexcel-test-01.xlsx", new FillItemBuilder()
                .add("userList", userList)
                .set("total", userList.size())
                .buildList());
    }

    @RequestMapping(value = "export4EasyExcel2")
    public void export4EasyExcel2() {
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
        ExcelUtils.export4EasyExcel(userList, User4EasyExcel.class, "easyexcel-test-02.xlsx");
    }

    @RequestMapping(value = "export4EasyPoi")
    public void export4EasyPoi() {
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
        ExcelUtils.export4EasyPoi("/excel/easypoi-test-01.xlsx", "easypoi-test-01.xlsx", model);
    }

    @RequestMapping(value = "export4Jxls")
    public void export4Jxls() {
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
        ExcelUtils.export4Jxls("/excel/jxls-test-01.xlsx", "test.xlsx", model);
    }

    @RequestMapping(value = "import4EasyExcel")
    @ResponseBody
    public void import4EasyExcel() {
        List<User4EasyExcel> dataList = ExcelUtils.import4EasyExcel(User4EasyExcel.class);
        dataList.forEach(System.out::println);
    }

    @RequestMapping(value = "import4EasyPoi")
    @ResponseBody
    public void import4EasyPoi() {
        List<User4EasyPoi> dataList = ExcelUtils.import4EasyPoi(User4EasyPoi.class);
        dataList.forEach(System.out::println);
    }
}
