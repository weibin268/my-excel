package com.zhuang.excel.easyexcel;

import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.metadata.fill.FillConfig;

public class HorizontalFillItem extends FillItem {
    public HorizontalFillItem(String name, Object data) {
        super(name, data, FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build());
    }
}
