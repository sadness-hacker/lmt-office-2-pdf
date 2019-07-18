package com.lmt.office.util;

import javax.servlet.http.HttpServletRequest;
import java.util.Properties;

/**
 * @description 获取文件上传路径工具类
 * @author bazhandao
 * @date 2019-07-17
 */
public class PropertiesUtils {

    /**
     * 默认配置文件
     */
    private static final String CONFIG_FILE = "application.properties";

    private static Properties properties;

    /**
     * 获取上传文件路径
     * @author bazhandao
     * @date 2019-07-17
     * @param request
     * @return
     */
    public static String getUploadPath(HttpServletRequest request) {
        try {
            if (properties == null) {
                properties = new Properties();
                properties.load(PropertiesUtils.class.getClassLoader().getResourceAsStream(CONFIG_FILE));
            }
            String uploadPath = properties.getProperty("file.upload.path");
            if (uploadPath == null || uploadPath.trim().equals("")) {
                uploadPath = request.getServletContext().getRealPath("/WEB-INF/upload");
            }
            return uploadPath;
        } catch (Exception e) {
            return request.getServletContext().getRealPath("/WEB-INF/upload");
        }
    }

    /**
     * 获取下载文件路径
     * @author bazhandao
     * @date 2019-07-17
     * @return
     */
    public static String getDownloadUrl() {
        try {
            if (properties == null) {
                properties = new Properties();
                properties.load(PropertiesUtils.class.getClassLoader().getResourceAsStream(CONFIG_FILE));
            }
            String downloadUrl = properties.getProperty("file.download.url");
            return downloadUrl;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取预览文件路径
     * @author bazhandao
     * @date 2019-07-17
     * @return
     */
    public static String getPreviewUrl() {
        try {
            if (properties == null) {
                properties = new Properties();
                properties.load(PropertiesUtils.class.getClassLoader().getResourceAsStream(CONFIG_FILE));
            }
            String downloadUrl = properties.getProperty("file.preview.url");
            return downloadUrl;
        } catch (Exception e) {
            return null;
        }
    }

}
