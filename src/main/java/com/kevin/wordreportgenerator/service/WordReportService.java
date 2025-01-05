package com.kevin.wordreportgenerator.service;

import cn.hutool.core.util.ObjectUtil;
import cn.hutool.extra.qrcode.QrCodeUtil;
import cn.hutool.extra.qrcode.QrConfig;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import com.kevin.wordreportgenerator.model.ReportGenerationRequest;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.awt.image.BufferedImage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
@Slf4j
public class WordReportService {

    @Value("${report.temp.dir:temp}")
    private String tempDir;

    public String generateReports(ReportGenerationRequest request) throws Exception {
        if (ObjectUtil.isEmpty(request.getData())) {
            throw new IllegalArgumentException("报告数据不能为空");
        }

        // 创建临时目录
        String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        String tempPath = tempDir + File.separator + timestamp;
        File tempFolder = new File(tempPath);
        if (!tempFolder.exists()) {
            tempFolder.mkdirs();
        }

        try {
            // 处理每条数据生成报告
            String reportName = request.getReportName();
            if(reportName == null) {
                reportName = "report";
            }
            for (int i = 0; i < request.getData().size(); i++) {
                Map<String, Object> data = request.getData().get(i);
                String outputPath = tempPath + File.separator + reportName + "_" + (i + 1) + ".docx";

                // 使用 poi-tl 配置模板（暂时不用）
                Configure config = Configure.builder()
                        .useSpringEL()
                        .build();

                // 创建模板对象
                XWPFTemplate template;
                if (request.getTemplateBase64() != null) {
                    template = XWPFTemplate.compile(new ByteArrayInputStream(
                            Base64ToFile.decodeDocx(request.getTemplateBase64())
                    ));
                } else {
                    template = XWPFTemplate.compile(request.getTemplatePath());
                }

                // 处理数据模型
                Map<String, Object> model = new HashMap<>(data);

                // 处理图片
                if (request.getImages() != null && request.getImages().size() > 0) {
                    processImages(model, request.getImages());
                }

                // 处理二维码
                if (request.getQrCodes() != null && request.getQrCodes().size() > 0) {
                    processQRCodes(model, data, request.getQrCodes());
                }

                // 渲染文档并保存
                template.render(model).writeToFile(outputPath);
            }

            // 打包所有报告
            String zipPath = tempPath + ".zip";
            zipFiles(tempPath, zipPath);

            System.out.println("================================>>>>>>>>>>> 报告生成完成!!!");
            return zipPath;
        } finally {
            // 清理临时文件
            FileUtils.deleteDirectory(tempFolder);
        }
    }

    private void processImages(Map<String, Object> model, List<Map<String, Object>> images) {
        for (Map<String, Object> image : images) {
            byte[] imageData = (byte[]) image.get("imageData");
            Map<String, Integer> size = (Map<String, Integer>) image.get("imageSize");
            int width = size.get("width");
            int height = size.get("height");
            String imageName = (String) image.get("imageName");
            model.put(imageName, Pictures.ofBytes(imageData)
                    .size(width, height)
                    .create());
        }
    }

    private void processQRCodes(Map<String, Object> model, Map<String, Object> data, List<Map<String, Object>> qrCodes) {
        for (Map<String, Object> qrCode : qrCodes) {
            Integer index = (Integer) qrCode.get("index");
            String name = (String) qrCode.get("name");
            Integer size = (Integer) qrCode.get("size");
            // 从数据中获取需要生成二维码的内容
            Object qrContent = data.get(String.valueOf(index));
            if (qrContent != null) {
                // 将二维码图片添加到模型中
                model.put(name, Pictures.ofBufferedImage(
                                generateQRCode(qrContent.toString()), PictureType.PNG)
                        .size(size, size).create());
            }
        }
    }

    public static BufferedImage generateQRCode(String text) {
        QrConfig config = new QrConfig(300, 300);
        config.setMargin(0);
        return QrCodeUtil.generate(
                text,
                config
        );
    }

    private void zipFiles(String sourcePath, String zipPath) throws IOException {
        File sourceDir = new File(sourcePath);
        FileOutputStream fos = new FileOutputStream(zipPath);
        ZipOutputStream zipOut = new ZipOutputStream(fos);

        for (File file : sourceDir.listFiles()) {
            FileInputStream fis = new FileInputStream(file);
            ZipEntry zipEntry = new ZipEntry(file.getName());
            zipOut.putNextEntry(zipEntry);

            byte[] bytes = new byte[1024];
            int length;
            while ((length = fis.read(bytes)) >= 0) {
                zipOut.write(bytes, 0, length);
            }
            fis.close();
        }

        zipOut.close();
        fos.close();
    }
}