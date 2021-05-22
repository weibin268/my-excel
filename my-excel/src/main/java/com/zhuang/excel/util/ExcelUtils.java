package com.zhuang.excel.util;

import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItem;
import com.zhuang.excel.easypoi.EasyPoiUtils;
import com.zhuang.excel.jxls.JxlsUtils;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.support.StandardMultipartHttpServletRequest;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

public class ExcelUtils {

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList) {
        export4EasyExcel(templateFilePath, fileName, fillItemList, getResponse());
    }

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

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName) {
        export4EasyExcel(data, head, fileName, getResponse());
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

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data) {
        export4EasyPoi(templateFilePath, fileName, data);
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        if (templateFilePath.startsWith("/")) {
            templateFilePath = templateFilePath.substring(1);
        }
        OutputStream outputStream = getOutputStream(fileName, response);
        EasyPoiUtils.export(templateFilePath, outputStream, data);
        try {
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data) {
        export4Jxls(templateFilePath, fileName, data);
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

    public static <T> List<T> import4EasyExcel(Class<T> head) {
        return import4EasyExcel(head, getRequest());
    }

    public static <T> List<T> import4EasyExcel(Class<T> head, HttpServletRequest request) {
        InputStream inputStream = getInputStream(request);
        List<T> dataList = EasyExcelUtils.readToList(inputStream, head);
        return dataList;
    }

    private static InputStream getInputStream(String templateFilePath) {
        InputStream inputStream = ExcelUtils.class.getResourceAsStream(templateFilePath);
        return inputStream;
    }

    private static InputStream getInputStream(HttpServletRequest request) {
        InputStream result = null;
        StandardMultipartHttpServletRequest multipartRequest = new StandardMultipartHttpServletRequest(request);
        for (Map.Entry<String, List<MultipartFile>> entry : multipartRequest.getMultiFileMap().entrySet()) {
            for (MultipartFile file : entry.getValue()) {
                try {
                    result = file.getInputStream();
                    break;
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
        return result;
    }

    private static OutputStream getOutputStream(String fileName, HttpServletResponse response) {
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

    /**
     * 获取当前会话的 request
     *
     * @return request
     */
    private static HttpServletRequest getRequest() {
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        if (servletRequestAttributes == null) {
            throw new RuntimeException("非Web上下文无法获取Request");
        }
        return servletRequestAttributes.getRequest();
    }

    /**
     * 获取当前会话的 response
     *
     * @return response
     */
    private static HttpServletResponse getResponse() {
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        if (servletRequestAttributes == null) {
            throw new RuntimeException("非Web上下文无法获取Request");
        }
        return servletRequestAttributes.getResponse();
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
