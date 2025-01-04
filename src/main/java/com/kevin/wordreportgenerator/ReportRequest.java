package com.kevin.wordreportgenerator;

import java.util.List;
import java.util.Map;

public class ReportRequest {
    private String templatePath; // 模板文件路径
    private String data; // 替换模板的动态数据

    // Getters and Setters
    public String getTemplatePath() {
        return templatePath;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }
}

