package com.zhuang.excel.util;

import com.zhuang.excel.easyexcel.EasyExcelUtils;
import com.zhuang.excel.easyexcel.FillItem;
import com.zhuang.excel.easypoi.EasyPoiUtils;
import com.zhuang.excel.hutool.HutoolExcelUtils;
import com.zhuang.excel.jxls.JxlsUtils;
import com.zhuang.excel.model.ExportOption;
import com.zhuang.excel.spring.SpringWebUtils;
import org.apache.poi.util.RecordFormatException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

public class ExcelUtils {

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList) {
        export4EasyExcel(templateFilePath, fileName, fillItemList, SpringWebUtils.getResponse());
    }

    public static void export4EasyExcel(String templateFilePath, String fileName, List<FillItem> fillItemList, HttpServletResponse response) {
        try (InputStream inputStream = SpringWebUtils.getFileInputStream(templateFilePath); OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            EasyExcelUtils.export(inputStream, outputStream, fillItemList);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName) {
        export4EasyExcel(data, head, fileName, SpringWebUtils.getResponse());
    }

    public static <T> void export4EasyExcel(List<T> data, Class<T> head, String fileName, HttpServletResponse response) {
        try (OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            EasyExcelUtils.export(outputStream, data, head);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data) {
        export4EasyPoi(templateFilePath, fileName, data, SpringWebUtils.getResponse());
    }

    public static void export4EasyPoi(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        if (templateFilePath.startsWith("/")) {
            templateFilePath = templateFilePath.substring(1);
        }
        try (OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            EasyPoiUtils.export(templateFilePath, outputStream, data);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export4EasyPoi(List<T> dataSet, Class<T> head, String fileName) {
        export4EasyPoi(dataSet, head, fileName, SpringWebUtils.getResponse());
    }

    public static <T> void export4EasyPoi(List<T> dataSet, Class<T> head, String fileName, HttpServletResponse response) {
        try (OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            EasyPoiUtils.export(outputStream, dataSet, head);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export4Hutool(List<T> dataSet, ExportOption exportOption, String fileName) {
        export4Hutool(dataSet, exportOption, fileName, SpringWebUtils.getResponse());
    }

    public static <T> void export4Hutool(List<T> dataSet, ExportOption exportOption, String fileName, HttpServletResponse response) {
        try (OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            HutoolExcelUtils.export(outputStream, dataSet, exportOption);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data) {
        export4Jxls(templateFilePath, fileName, data, SpringWebUtils.getResponse());
    }

    public static void export4Jxls(String templateFilePath, String fileName, Map<String, Object> data, HttpServletResponse response) {
        try (InputStream inputStream = SpringWebUtils.getFileInputStream(templateFilePath); OutputStream outputStream = SpringWebUtils.getFileOutputStream(fileName, response)) {
            JxlsUtils.export(inputStream, outputStream, data);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> import4EasyExcel(Class<T> head) {
        return import4EasyExcel(head, SpringWebUtils.getRequest());
    }

    public static <T> List<T> import4EasyExcel(Class<T> head, HttpServletRequest request) {
        InputStream inputStream = SpringWebUtils.getFileInputStream(request);
        return EasyExcelUtils.readToList(inputStream, head);
    }

    public static <T> List<T> import4EasyPoi(Class<T> pojoClass) {
        return import4EasyPoi(pojoClass, SpringWebUtils.getRequest());
    }

    public static <T> List<T> import4EasyPoi(Class<T> pojoClass, HttpServletRequest request) {
        InputStream inputStream = SpringWebUtils.getFileInputStream(request);
        return EasyPoiUtils.readToList(inputStream, pojoClass);
    }

    public static void downloadTemplate(String templateFilePath) {
        String fileName = Paths.get(templateFilePath).getFileName().toString();
        downloadTemplate(templateFilePath, fileName, SpringWebUtils.getResponse());
    }

    public static void downloadTemplate(String templateFilePath, String fileName) {
        downloadTemplate(templateFilePath, fileName, SpringWebUtils.getResponse());
    }

    public static void downloadTemplate(String templateFilePath, String fileName, HttpServletResponse response) {
        try {
            SpringWebUtils.toFileResponse(response, fileName);
            OutputStream outputStream = response.getOutputStream();
            InputStream inputStream = ExcelUtils.class.getResourceAsStream(templateFilePath);
            // 没设置ContentLength，打开excel会提示需要修复文件
            response.setContentLength(inputStream.available());
            copy(inputStream, outputStream);
            inputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void copy(InputStream inp, OutputStream out) throws IOException {
        byte[] buff = new byte[4096];
        int count;
        while ((count = inp.read(buff)) != -1) {
            if (count > 0) {
                out.write(buff, 0, count);
            }
        }
    }
}
