package com.kevin.wordreportgenerator;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

public class TemplateService {
    private static final String TEMPLATE_RECORD_FILE = "data/templates.json";
    private static final String TEMPLATE_DIR = "templates/";
    private static final ObjectMapper objectMapper = new ObjectMapper();

    public static void addTemplateRecord(String name, String path) throws IOException {
        List<Map<String, String>> templates = getAllTemplates();
        Map<String, String> newTemplate = new HashMap<>();
        newTemplate.put("name", name);
        newTemplate.put("path", path);
        templates.add(newTemplate);

        Files.write(Paths.get(TEMPLATE_RECORD_FILE), objectMapper.writeValueAsBytes(templates));
    }

    public static List<Map<String, String>> getAllTemplates() throws IOException {
        Path filePath = Paths.get(TEMPLATE_RECORD_FILE);
        if (!Files.exists(filePath)) {
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, "[]".getBytes());
        }
        return objectMapper.readValue(Files.readAllBytes(filePath), List.class);
    }

    // 删除模板文件和记录
    public static boolean deleteTemplate(String name) throws IOException {
        List<Map<String, String>> templates = getAllTemplates();

        // 查找并删除记录
        Optional<Map<String, String>> templateToDelete = templates.stream()
                .filter(t -> t.get("name").equals(name))
                .findFirst();

        if (templateToDelete.isPresent()) {
            // 删除文件
            Path templatePath = Paths.get(TEMPLATE_DIR + name);
            Files.deleteIfExists(templatePath);

            // 更新JSON记录
            List<Map<String, String>> updatedTemplates = templates.stream()
                    .filter(t -> !t.get("name").equals(name))
                    .collect(Collectors.toList());

            Files.write(Paths.get(TEMPLATE_RECORD_FILE), objectMapper.writeValueAsBytes(updatedTemplates));
            return true;
        }

        return false;
    }

}