package com.zhuang.excel.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
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
        try {
            OutputStream outputStream = new FileOutputStream(outputFilePath);
            export(inputFilePath, outputStream, model);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(OutputStream outputStream, Collection<T> dataSet, Class<T> pojoClass) {
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(null,null, ExcelType.XSSF), pojoClass, dataSet);
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(String outputFilePath, Collection<T> dataSet, Class<T> pojoClass) {
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outputFilePath);
            export(outputStream, dataSet, pojoClass);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }
}
