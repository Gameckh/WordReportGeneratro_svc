package com.kevin.wordreportgenerator.model;

import lombok.Data;
import java.util.List;
import java.util.Map;

@Data
public class ReportGenerationRequest {
    private List<Map<String, Object>> data; // JSON数组数据

    private String templateBase64; // Word模板内容（Base64）
    private byte[] templateContent; // Word模板内容（二进制）
    private String templatePath; // Word模板路径

    // 图片集合
    /*
    * imageData -> byte[]
    * imageName -> String
    * imageSize -> Map<String, Integer>
    * */
    private List<Map<String, Object>> images;

    // 二维码集合
    /*
        index -> Integer
        name -> String
        size -> Integer
    * */
    private List<Map<String, Object>> qrCodes;
    private String reportName; // 报告名称前缀（后缀是顺序）
    private String returnName; // 返回的zip包名称
}