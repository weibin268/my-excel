package com.zhuang.excel.util;

import org.junit.Test;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.*;

public class JxlsUtilsTest {

    @Test
    public void export() {
        Map<String, Object> model = new HashMap<>();
        model.put("name","zwb");
        model.put("age",100);
        model.put("date",new Date());
        JxlsUtils.export(getClass().getResource("/excel/jxls-test-01.xlsx").getPath(),"D:\\temp\\jxls-test-01.xlsx",model);
    }
}