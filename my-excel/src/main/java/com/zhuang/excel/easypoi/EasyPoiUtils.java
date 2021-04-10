package com.zhuang.excel.easypoi;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.zhuang.excel.easyexcel.FillItem;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
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
        try {
            OutputStream outputStream = new FileOutputStream(outputFilePath);
            export(inputFilePath, outputStream, model);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
