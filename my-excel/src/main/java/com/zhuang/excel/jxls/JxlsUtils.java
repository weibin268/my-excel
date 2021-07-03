package com.zhuang.excel.jxls;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class JxlsUtils {

    public static final String FUNCTION_MAP_KEY = "FUNCTION_MAP";

    public static void export(InputStream inputStream, OutputStream outputStream, Map<String, Object> model) {
        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        //重新计算公式
        jxlsHelper.setEvaluateFormulas(true);
        Transformer transformer = jxlsHelper.createTransformer(inputStream, outputStream);
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
        Map<String, Object> functionMap = getFunctionMap(model);
        JexlEngine jexlEngine = new JexlBuilder().namespaces(functionMap).create();
        evaluator.setJexlEngine(jexlEngine);
        try {
            jxlsHelper.processTemplate(context, transformer);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void export(String inputFilePath, String outputFilePath, Map<String, Object> model) {
        try (InputStream inputStream = new FileInputStream(inputFilePath); OutputStream outputStream = new FileOutputStream(outputFilePath)) {
            export(inputStream, outputStream, model);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static Map<String, Object> getFunctionMap(Map<String, Object> model) {
        Map<String, Object> functionMap = new HashMap<>();
        functionMap.put("utils", new JxlsUtils());
        if (model != null) {
            Object ojbFunctionMap = model.get(FUNCTION_MAP_KEY);
            if (ojbFunctionMap != null) {
                if (ojbFunctionMap instanceof Map) {
                    Map<String, Object> tempFunctionMap = (Map<String, Object>) ojbFunctionMap;
                    tempFunctionMap.forEach((c, value) -> {
                        functionMap.put(c, value);
                    });
                }
            }
        }
        return functionMap;
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

}
