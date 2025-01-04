package com.kevin.wordreportgenerator;

import com.kevin.wordreportgenerator.ReportService;
import com.kevin.wordreportgenerator.ReportRequest;
import org.apache.commons.io.IOUtils;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.File;

@RestController
@RequestMapping("/api/reports")
public class ReportController {

    @PostMapping("/generate")
    public ResponseEntity<?> generateReports(@RequestBody ReportRequest request) {
        try {
            // 调用 ReportService 生成 ZIP 文件
            String zipFilePath = ReportService.generateReports(request);
            File zipFile = new File(zipFilePath);

            // 读取生成的 ZIP 文件并返回
            FileInputStream fileInputStream = new FileInputStream(zipFile);
            byte[] zipContent = IOUtils.toByteArray(fileInputStream);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + zipFile.getName())
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(zipContent);

        } catch (FileNotFoundException e) {
            return ResponseEntity.status(404).body("ZIP file not found");
        } catch (IOException e) {
            return ResponseEntity.status(500).body("Error generating reports");
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Internal server error");
        }
    }
}
