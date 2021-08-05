package com.zhuang.excel.jxls;

import java.text.SimpleDateFormat;
import java.util.Date;

public class CommonUtils {

    public String formatDate(Date date, String format) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
            return simpleDateFormat.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }


}
