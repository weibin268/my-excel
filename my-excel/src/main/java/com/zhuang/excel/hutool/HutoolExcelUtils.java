package com.zhuang.excel.hutool;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.zhuang.excel.model.ExportOption;
import org.apache.poi.ss.usermodel.Font;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class HutoolExcelUtils {

    public static <T> void export(OutputStream outputStream, Collection<T> dataSet, ExportOption exportOption) {
        List<String> fieldNameList = exportOption.getColumnList().stream().map(ExportOption.Column::getFieldName).collect(Collectors.toList());
        List<Map<String, Object>> dataMap = dataSet.stream().map(item -> {
            Map<String, Object> map = BeanUtil.beanToMap(item);
            Map<String, Object> newMap = new LinkedHashMap<>();
            for (String fieldName : fieldNameList) {
                if (map.containsKey(fieldName)) {
                    newMap.put(fieldName, map.get(fieldName));
                }
            }
            return newMap;
        }).collect(Collectors.toList());
        export(outputStream, dataMap, exportOption);
    }

    public static <T> void export(OutputStream outputStream, List<Map<String, Object>> dataMap, ExportOption exportOption) {
        ExcelWriter writer = ExcelUtil.getBigWriter();
        final Font headFont = writer.getWorkbook().createFont();
        headFont.setBold(true);
        headFont.setFontName("宋体");
        headFont.setFontHeightInPoints((short) 14);
        writer.getHeadCellStyle().setFont(headFont);
        for (int i = 0; i < exportOption.getColumnList().size(); i++) {
            ExportOption.Column column = exportOption.getColumnList().get(i);
            writer.addHeaderAlias(column.getFieldName(), column.getHeadName());
            writer.setColumnWidth(i, column.getWidth());
        }
        writer.write(dataMap).flush(outputStream);
    }

    public static <T> void export(String outputFilePath, List<Map<String, Object>> dataMap, ExportOption exportOption) {
        try (OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            export(outputStream, dataMap, exportOption);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> void export(String outputFilePath, Collection<T> dataSet, ExportOption exportOption) {
        try (OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            export(outputStream, dataSet, exportOption);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
