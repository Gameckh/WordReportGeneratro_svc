package com.kevin.wordreportgenerator;

import jakarta.annotation.PostConstruct;
import org.springframework.stereotype.Component;

import java.io.File;

@Component
public class DirectoryInitializer {

    private final String[] dirs = {"templates/", "data_json/", "generated_reports/"};

    @PostConstruct
    public void init() {
        for (String dir : dirs) {
            File directory = new File(dir);
            if (!directory.exists()) {
                directory.mkdirs();
            }
        }
    }
}
