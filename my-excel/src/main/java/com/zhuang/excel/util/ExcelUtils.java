package com.zhuang.excel.util;

import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItem;
import com.zhuang.excel.easypoi.EasyPoiUtils;
import com.zhuang.excel.jxls.JxlsUtils;
import com.zhuang.excel.spring.SpringMvcUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLDecoder;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

public class ExcelUtils {

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList) {
        export4EasyExcel(templateFilePath, fileName, fillItemList, SpringMvcUtils.getResponse());
    }

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList, HttpServletResponse response) {
        try (InputStream inputStream = SpringMvcUtils.getFileInputStream(templateFilePath); OutputStream outputStream = SpringMvcUtils.getFileOutputStream(fileName, response)) {
            EasyExcelUtils.export(inputStream, outputStream, fillItemList);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName) {
        export4EasyExcel(data, head, fileName, SpringMvcUtils.getResponse());
    }

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName, HttpServletResponse response) {
        try (OutputStream outputStream = SpringMvcUtils.getFileOutputStream(fileName, response)) {
            EasyExcelUtils.export(outputStream, data, head);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data) {
        export4EasyPoi(templateFilePath, fileName, data, SpringMvcUtils.getResponse());
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        if (templateFilePath.startsWith("/")) {
            templateFilePath = templateFilePath.substring(1);
        }
        try (OutputStream outputStream = SpringMvcUtils.getFileOutputStream(fileName, response)) {
            EasyPoiUtils.export(templateFilePath, outputStream, data);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export4EasyPoi(List<T> dataSet, Class<T> head, String fileName) {
        export4EasyPoi(dataSet, head, fileName, SpringMvcUtils.getResponse());
    }

    public static <T> void export4EasyPoi(List<T> dataSet, Class<T> head, String fileName, HttpServletResponse response) {
        try (OutputStream outputStream = SpringMvcUtils.getFileOutputStream(fileName, response)) {
            EasyPoiUtils.export(outputStream, dataSet, head);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data) {
        export4Jxls(templateFilePath, fileName, data, SpringMvcUtils.getResponse());
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        try (InputStream inputStream = SpringMvcUtils.getFileInputStream(templateFilePath); OutputStream outputStream = SpringMvcUtils.getFileOutputStream(fileName, response)) {
            JxlsUtils.export(inputStream, outputStream, data);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> import4EasyExcel(Class<T> head) {
        return import4EasyExcel(head, SpringMvcUtils.getRequest());
    }

    public static <T> List<T> import4EasyExcel(Class<T> head, HttpServletRequest request) {
        InputStream inputStream = SpringMvcUtils.getFileInputStream(request);
        return EasyExcelUtils.readToList(inputStream, head);
    }

    public static <T> List<T> import4EasyPoi(Class<T> pojoClass) {
        return import4EasyPoi(pojoClass, SpringMvcUtils.getRequest());
    }

    public static <T> List<T> import4EasyPoi(Class<T> pojoClass, HttpServletRequest request) {
        InputStream inputStream = SpringMvcUtils.getFileInputStream(request);
        return EasyPoiUtils.readToList(inputStream, pojoClass);
    }

    public static void downloadTemplate(String templateFilePath) {
        String fileName = Paths.get(templateFilePath).getFileName().toString();
        downloadTemplate(templateFilePath, fileName, SpringMvcUtils.getResponse());
    }

    public static void downloadTemplate(String templateFilePath, String fileName) {
        downloadTemplate(templateFilePath, fileName, SpringMvcUtils.getResponse());
    }

    public static void downloadTemplate(String templateFilePath, String fileName, HttpServletResponse response) {
        try {
            SpringMvcUtils.toFileResponse(response, fileName);
            OutputStream outputStream = response.getOutputStream();
            String path = ExcelUtils.class.getResource(templateFilePath).getPath();
            path = URLDecoder.decode(path, SpringMvcUtils.DEFAULT_CHARSET);
            Files.copy(new File(path).toPath(), outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
