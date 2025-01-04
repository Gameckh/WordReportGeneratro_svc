package com.kevin.wordreportgenerator;

import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.*;

@RestController
@RequestMapping("/api/templates")
public class TemplateController {

    private static final String TEMPLATE_DIR = "templates/";

    // 上传模板
    @PostMapping("/upload")
    public ResponseEntity<?> uploadTemplate(@RequestParam("file") MultipartFile file) {
        try {
            // 保存文件
            String fileName = file.getOriginalFilename();
            Path path = Paths.get(TEMPLATE_DIR + fileName);
            Files.createDirectories(path.getParent());
            Files.copy(file.getInputStream(), path, StandardCopyOption.REPLACE_EXISTING);

            // 更新JSON记录
            TemplateService.addTemplateRecord(fileName, path.toString());
            return ResponseEntity.ok("Template uploaded successfully");
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Upload failed");
        }
    }

    // 获取所有模板
    @GetMapping("/list")
    public ResponseEntity<?> listTemplates() throws IOException {
        return ResponseEntity.ok(TemplateService.getAllTemplates());
    }

    @DeleteMapping("/delete/{name}")
    public ResponseEntity<?> deleteTemplate(@PathVariable String name) {
        try {
            if (TemplateService.deleteTemplate(name)) {
                return ResponseEntity.ok("Template deleted");
            } else {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body("Template not found");
            }
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Delete failed");
        }
    }

    @GetMapping("/download/{name}")
    public ResponseEntity<?> downloadTemplate(@PathVariable String name) {
        try {
            Path filePath = Paths.get(TEMPLATE_DIR + name);
            if (Files.exists(filePath)) {
                File file = filePath.toFile();

                return ResponseEntity.ok()
                        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + file.getName())
                        .header(HttpHeaders.CONTENT_TYPE, "application/octet-stream")
                        .body(new InputStreamResource(new FileInputStream(file)));
            } else {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body("Template not found");
            }
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to download template");
        }
    }

    @PostMapping("/generate")
    public ResponseEntity<?> generateReports(@RequestBody ReportRequest request) {
        try {
            String zipPath = ReportService.generateReports(request);
            File zipFile = new File(zipPath);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + zipFile.getName())
                    .body(new InputStreamResource(new FileInputStream(zipFile)));
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to generate reports");
        }
    }
}
