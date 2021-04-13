package com.zhuang.excel.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillWrapper;

import java.io.*;
import java.util.ArrayList;
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
            outputStream.close();
            inputStream.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(OutputStream outputStream, List<T> data, Class<T> head) {
        EasyExcel.write(outputStream, head).sheet().doWrite(data);
    }

    public static <T> void export(String outputFilePath, List<T> data, Class<T> head) {
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outputFilePath);
            EasyExcel.write(outputStream, head).sheet().doWrite(data);
            outputStream.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> readToList(String inputFilePath, Class head) {
        try {
            InputStream inputStream = new FileInputStream(inputFilePath);
            return readToList(inputStream, head);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> readToList(InputStream inputStream, Class head) {
        List<T> result = new ArrayList<>();
        EasyExcel.read(inputStream, head, new AnalysisEventListener<T>() {
            @Override
            public void invoke(T data, AnalysisContext context) {
                result.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
            }
        }).sheet().doRead();
        return result;
    }
}
