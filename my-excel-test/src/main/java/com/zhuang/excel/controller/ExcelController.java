package com.zhuang.excel.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
@RequestMapping(value = "/excel")
public class ExcelController {

    @ResponseBody
    @GetMapping(value = "test")
    public String test() {
        return "{test:123}";
    }

}
