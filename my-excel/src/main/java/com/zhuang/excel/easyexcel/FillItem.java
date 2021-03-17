package com.zhuang.excel.easyexcel;

import com.alibaba.excel.write.metadata.fill.FillConfig;

public class FillItem {

    private String name;
    private Object data;
    private FillConfig fillConfig;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public FillConfig getFillConfig() {
        return fillConfig;
    }

    public void setFillConfig(FillConfig fillConfig) {
        this.fillConfig = fillConfig;
    }

    public FillItem(String name, Object data) {
        this.name = name;
        this.data = data;
        this.fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
    }

    public FillItem(String name, Object data, FillConfig fillConfig) {
        this.setName(name);
        this.data = data;
        this.setFillConfig(fillConfig);
    }
}
