package com.kevin.wordreportgenerator.service;
import java.util.Base64;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Base64ToFile {

    private static final Logger LOGGER = Logger.getLogger(Base64ToFile.class.getName());

    /**
     * 将 Base64 编码的 docx 文件解码为字节数组
     *
     * @param base64String Base64 编码的 docx 文件内容，可以包含 Data URL 前缀
     * @return 解码后的字节数组
     * @throws IllegalArgumentException 如果传入的参数为空或无法正常解码
     */
    public static byte[] decodeDocx(String base64String) {
        // 1. 检查输入是否为空
        if (Objects.isNull(base64String) || base64String.trim().isEmpty()) {
            String msg = "Base64 编码字符串不能为空。";
            LOGGER.log(Level.WARNING, msg);
            throw new IllegalArgumentException(msg);
        }

        // 2. 移除 Data URL 前缀（如果存在）
        String pureBase64;
        int commaIndex = base64String.indexOf(',');
        if (commaIndex != -1) {
            pureBase64 = base64String.substring(commaIndex + 1);
            LOGGER.log(Level.INFO, "检测到 Data URL 前缀，已移除前缀部分。");
        } else {
            pureBase64 = base64String;
        }

        // 3. 尝试进行 Base64 解码，捕获潜在的异常
        try {
            return Base64.getDecoder().decode(pureBase64);
        } catch (IllegalArgumentException e) {
            String msg = "Base64 编码的 docx 文件内容非法或无法正常解码。";
            LOGGER.log(Level.SEVERE, msg, e);
            throw new IllegalArgumentException(msg, e);
        }
    }
}
