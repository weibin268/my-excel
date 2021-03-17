package com.zhuang.excel.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillWrapper;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

public class EasyExcelUtils {

    public static void export(InputStream inputStream, OutputStream outputStream, List<FillItem> fillItemList) {
        ExcelWriter excelWriter = EasyExcel.write(outputStream).withTemplate(inputStream).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        for (FillItem fillItem : fillItemList) {
            if (fillItem.getData() instanceof Collection) {
                excelWriter.fill(new FillWrapper(fillItem.getName(), (Collection) fillItem.getData()), fillItem.getFillConfig(), writeSheet);
            } else {
                excelWriter.fill(fillItem.getData(), fillItem.getFillConfig(), writeSheet);
            }
        }
        excelWriter.finish();
    }

    public static void export(String inputFilePath, String outputFilePath, List<FillItem> fillItemList) {
        try {
            InputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);
            export(inputStream, outputStream, fillItemList);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
