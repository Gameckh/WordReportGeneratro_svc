package com.kevin.wordreportgenerator.controller;

import com.kevin.wordreportgenerator.model.ReportGenerationRequest;
import com.kevin.wordreportgenerator.service.WordReportService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;

@RestController
@RequestMapping("/api/word-report")
@Slf4j
public class WordReportController {

    @Autowired
    private WordReportService wordReportService;

    @PostMapping("/generate")
    public ResponseEntity<?> generateReports(@RequestBody ReportGenerationRequest request) {
        try {
            // 验证请求参数
            if (request.getData() == null || request.getData().isEmpty()) {
                return ResponseEntity.badRequest().body("报告数据不能为空");
            }
            if (request.getTemplateBase64() == null
                    && (request.getTemplatePath() == null || request.getTemplatePath().isEmpty())) {
                return ResponseEntity.badRequest().body("模板内容或路径不能为空");
            }

            String zipPath = wordReportService.generateReports(request);
            File zipFile = new File(zipPath);

            String returnName = request.getReturnName();
            if(returnName == null) {
                returnName = "reports";
            }

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + returnName + ".zip")
                    .header(HttpHeaders.CONTENT_TYPE, "application/zip")
                    .body(new InputStreamResource(new FileInputStream(zipFile)));
        } catch (Exception e) {
            log.error("生成报告失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("生成报告失败: " + e.getMessage());
        }
    }
}