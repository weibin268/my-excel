package com.zhuang.excel.util;

import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItem;
import com.zhuang.excel.easypoi.EasyPoiUtils;
import com.zhuang.excel.jxls.JxlsUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class ExcelUtils {

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList, HttpServletResponse response) {
        InputStream inputStream = getInputStream(templateFilePath);
        OutputStream outputStream = getOutputStream(fileName, response);
        EasyExcelUtils.export(inputStream, outputStream, fillItemList);
        try {
            outputStream.close();
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName, HttpServletResponse response) {
        OutputStream outputStream = getOutputStream(fileName, response);
        EasyExcelUtils.export(outputStream, data, head);
        try {
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        OutputStream outputStream = getOutputStream(fileName, response);
        EasyPoiUtils.export(templateFilePath, outputStream, data);
        try {
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        InputStream inputStream = getInputStream(templateFilePath);
        OutputStream outputStream = getOutputStream(fileName, response);
        JxlsUtils.export(inputStream, outputStream, data);
        try {
            outputStream.close();
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static InputStream getInputStream(String templateFilePath) {
        InputStream inputStream = ExcelUtils.class.getResourceAsStream(templateFilePath);
        return inputStream;
    }

    public static OutputStream getOutputStream(String fileName, HttpServletResponse response) {
        try {
            fileName = resolveFileNameEncoding(fileName);
            //fileName=URLEncoder.encode(fileName,"utf-8");//IE
            response.setHeader("content-disposition", "attachment;filename=" + fileName);
            ServletOutputStream outputStream = response.getOutputStream();
            return outputStream;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
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
