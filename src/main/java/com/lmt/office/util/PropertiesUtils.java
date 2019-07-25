package com.lmt.office.util;

import lombok.extern.slf4j.Slf4j;

import javax.servlet.http.HttpServletRequest;
import java.util.Properties;

/**
 * @description 获取文件上传路径工具类
 * @author bazhandao
 * @date 2019-07-17
 */
@Slf4j
public class PropertiesUtils {

    /**
     * 默认配置文件
     */
    private static final String CONFIG_FILE = "application.properties";

    private static Properties properties = getProperties();

    private static Properties getProperties() {
        try {
            Properties properties = new Properties();
            properties.load(PropertiesUtils.class.getClassLoader().getResourceAsStream(CONFIG_FILE));
            return properties;
        } catch (Exception e) {
            log.error("加载配置文件出错!!!file={},{}", CONFIG_FILE, e);
        }
        return null;
    }

    /**
     * 获取上传文件路径
     * @author bazhandao
     * @date 2019-07-17
     * @param request
     * @return
     */
    public static String getUploadPath(HttpServletRequest request) {
        try {
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
            String downloadUrl = properties.getProperty("file.preview.url");
            return downloadUrl;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取LibreOffice命令
     * @author bazhandao
     * @date 2019-07-25
     * @return
     */
    public static String getLibreOfficeCmd() {
        try {
            String downloadUrl = properties.getProperty("libreoffice.cmd");
            return downloadUrl;
        } catch (Exception e) {
            return null;
        }
    }

}
