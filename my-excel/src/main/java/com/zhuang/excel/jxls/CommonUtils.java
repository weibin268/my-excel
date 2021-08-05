package com.zhuang.excel.jxls;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CommonUtils {

    public String formatDate(Date date) {
        return formatDate(date, "yyyy-MM-dd HH:mm:ss");
    }

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

    public String toString(BigDecimal num) {
        return toString(num, false);
    }

    public String toString(BigDecimal num, boolean stripTrailingZeros) {
        if (num == null) return null;
        if (stripTrailingZeros) {
            num = num.stripTrailingZeros();
        }
        return num.toPlainString();
    }
}
