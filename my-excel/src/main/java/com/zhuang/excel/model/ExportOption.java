package com.zhuang.excel.model;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class ExportOption {

    private List<Column> columnList = new ArrayList<>();

    private Integer defaultColumnWidth;

    public ExportOption addColumn(String fieldName, String headName, Integer width) {
        Column column = new Column();
        column.setFieldName(fieldName);
        column.setHeadName(headName);
        column.setWidth(width);
        columnList.add(column);
        return this;
    }

    public ExportOption addColumn(String fieldName, String headName) {
        return addColumn(fieldName, headName, defaultColumnWidth);
    }

    @Data
    public static class Column {
        private String fieldName;
        private String headName;
        private Integer width;
    }
}
