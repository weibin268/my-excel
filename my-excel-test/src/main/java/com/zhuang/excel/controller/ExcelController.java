package com.zhuang.excel.controller;

import com.zhuang.excel.jxls.JxlsUtils;
import com.zhuang.excel.model.User;
import com.zhuang.excel.util.ExcelUtils;
import org.apache.commons.collections.map.HashedMap;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.*;

@Controller
@RequestMapping(value = "/excel")
public class ExcelController {

    @ResponseBody
    @GetMapping(value = "test")
    public String test() {
        return "{test:123}";
    }

    @RequestMapping(value = "export")
    public void export(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String templateFilePath = "/excel/jxls-test-01.xlsx";
        String fileName = "test.xlsx";
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
        ExcelUtils.export(templateFilePath, fileName, model, response);
    }

}
