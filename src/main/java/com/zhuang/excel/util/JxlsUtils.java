package com.zhuang.excel.util;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class JxlsUtils {

    public static void export(InputStream inputStream, OutputStream outputStream, Map<String, Object> model) throws IOException {
        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer = TransformerFactory.createTransformer(inputStream, outputStream);
        //获得配置
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
        //设置静默模式，不报警告
        //evaluator.getJexlEngine().setSilent(true);
        //函数强制，自定义功能
        Map<String, Object> functionMap = new HashMap<>();
        functionMap.put("utils", new JxlsUtils());
        JexlEngine customJexlEngine = new JexlBuilder().namespaces(functionMap).create();
        evaluator.setJexlEngine(customJexlEngine);
        //必须要这个，否者表格函数统计会错乱
        jxlsHelper.setUseFastFormulaProcessor(false);
        jxlsHelper.processTemplate(context, transformer);
    }

    public static void export(String inputFilePath, String outputFilePath, Map<String, Object> model) {
        try {
            InputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);
            export(inputStream, outputStream, model);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public String formatDate(Date date, String format) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(format);
            return dateFmt.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

}
