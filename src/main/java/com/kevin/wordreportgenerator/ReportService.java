package com.kevin.wordreportgenerator;

import cn.hutool.extra.qrcode.QrCodeUtil;
import cn.hutool.extra.qrcode.QrConfig;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.zip.*;

public class ReportService {
    public static String generateReports(ReportRequest request) throws Exception {
        String outputDir = "reports/";
        String zipFilePath = "reports/reports.zip";
        Files.createDirectories(Paths.get(outputDir));

        ObjectMapper mapper = new ObjectMapper();
        List<List<String>> data = mapper.readValue(request.getData(), List.class);

        List<Map<String, Object>> maps = new ArrayList<>();
        for (int i = 0; i < data.size(); i++) {
            List<String> item = data.get(i);
            Map<String, Object> map = new HashMap<>();
            for (int j = 0; j < item.size(); j++) {
                map.put(String.valueOf(j), item.get(j));
            }
            map.put("QR", Pictures.ofBufferedImage(generateQRCode(item.get(1)), PictureType.PNG)
                    .size(47, 47).create());
            maps.add(map);
        }

        for (int i = 0; i < maps.size(); i++) {
            String reportName = "report_" + (i+1) + ".docx";
            Map<String, Object> map = maps.get(i);
            XWPFTemplate.compile(request.getTemplatePath()).render(map).writeToFile(outputDir + reportName);
        }

        // 打包ZIP
        ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(zipFilePath));
        for (File file : Objects.requireNonNull(new File(outputDir).listFiles())) {
            FileInputStream fis = new FileInputStream(file);
            ZipEntry zipEntry = new ZipEntry(file.getName());
            zipOut.putNextEntry(zipEntry);
            IOUtils.copy(fis, zipOut);
            fis.close();
        }
        zipOut.close();
        return zipFilePath;
    }

    public static BufferedImage generateQRCode(String text) throws IOException {
        QrConfig config = new QrConfig(300, 300);
        config.setMargin(0);
        BufferedImage bufferedImage = QrCodeUtil.generate(
                text,
                config
        );
        return bufferedImage;
    }
}
