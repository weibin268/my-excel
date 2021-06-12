package com.zhuang.excel.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

public class EasyPoiUtils {

    public static void export(String inputFilePath, OutputStream outputStream, Map<String, Object> model) {
        TemplateExportParams templateExportParams = new TemplateExportParams(inputFilePath);
        Workbook workbook = ExcelExportUtil.exportExcel(templateExportParams, model);
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export(String inputFilePath, String outputFilePath, Map<String, Object> model) {
        try (OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            export(inputFilePath, outputStream, model);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(OutputStream outputStream, Collection<T> dataSet, Class<T> pojoClass) {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(null, null, ExcelType.XSSF), pojoClass, dataSet);
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(String outputFilePath, Collection<T> dataSet, Class<T> pojoClass) {
        try (OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            export(outputStream, dataSet, pojoClass);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> readToList(String inputFilePath, Class<?> pojoClass) {
        try (InputStream inputStream = new FileInputStream(inputFilePath)) {
            return readToList(inputStream, pojoClass);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> readToList(InputStream inputStream, Class<?> pojoClass) {
        List<T> result;
        ImportParams importParams = new ImportParams();
        try {
            result = ExcelImportUtil.importExcel(inputStream, pojoClass, importParams);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
