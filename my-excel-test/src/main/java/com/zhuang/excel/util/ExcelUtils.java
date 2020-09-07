package com.zhuang.excel.util;

import com.zhuang.excel.jxls.JxlsUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.util.Map;

public class ExcelUtils {

    public static void export(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) throws IOException {
        fileName = resolveFileNameEncoding(fileName);
        //fileName=URLEncoder.encode(fileName,"utf-8");//IE
        response.setHeader("content-disposition", "attachment;filename=" + fileName);
        ServletOutputStream outputStream = response.getOutputStream();
        InputStream inputStream = ExcelUtils.class.getResourceAsStream(templateFilePath);
        JxlsUtils.export(inputStream, outputStream, data);
        outputStream.close();
        inputStream.close();
    }

    private static String resolveFileNameEncoding(String fileName) {
        String result = null;
        try {
            result = new String(fileName.getBytes("utf-8"), "ISO8859-1");//chrome,firefox
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
