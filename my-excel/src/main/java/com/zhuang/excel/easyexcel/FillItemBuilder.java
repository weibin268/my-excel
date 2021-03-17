package com.zhuang.excel.easyexcel;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class FillItemBuilder {

    private List<FillItem> fillItemList = new ArrayList<>();

    public FillItemBuilder add(Object data) {
        fillItemList.add(new FillItem(null, data));
        return this;
    }

    public FillItemBuilder add(String name, Collection data) {
        fillItemList.add(new FillItem(name, data));
        return this;
    }

    public FillItemBuilder add4Horizontal(String name, Collection collection) {
        fillItemList.add(new HorizontalFillItem(name, collection));
        return this;
    }

    List<FillItem> buildList() {
        return fillItemList;
    }

}
