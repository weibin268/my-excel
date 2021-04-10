package com.zhuang.excel.easyexcel;

import java.util.*;

public class FillItemBuilder {

    private List<FillItem> fillItemList = new ArrayList<>();
    private Map<String, Object> map = new HashMap<String, Object>();


    public FillItemBuilder add(String name, Object data) {
        fillItemList.add(new FillItem(name, data));
        return this;
    }

    public FillItemBuilder add4Horizontal(String name, Object data) {
        fillItemList.add(new HorizontalFillItem(name, data));
        return this;
    }

    public FillItemBuilder set(String key, Object value) {
        map.put(key, value);
        return this;
    }

    public List<FillItem> buildList() {
        if (map.size() > 0) {
            add(null, map);
        }
        return fillItemList;
    }

}
