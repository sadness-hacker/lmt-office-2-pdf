package com.lmt.office.util;

import lombok.extern.slf4j.Slf4j;

/**
 * @description Linux下通过LibreOffice转换为pdf
 *
 * @author bazhandao
 * @date 2019/7/25 16:16
 * @since JDK1.8
 */
@Slf4j
public class LibreOfficeUtils {

    /**
     * word/excel/ppt转pdf工具类
     * @author bazhandao
     * @date 2019-07-17
     * @param filePath     源文件路径
     * @param pdfFilePath  pdf文件路径
     */
    public static void file2pdf(String filePath, String pdfFilePath) {
        int dot = filePath.lastIndexOf('.');
        String suffix = filePath.substring(dot).toLowerCase();
        if (".doc".equals(suffix) || ".docx".equals(suffix)
            || ".ppt".equals(suffix) || ".pptx".equals(suffix)
            || ".xls".equals(suffix) || ".xlsx".equals(suffix)) {
            try {
                dot = pdfFilePath.lastIndexOf('/');
                String cmd = PropertiesUtils.getLibreOfficeCmd();
                cmd = cmd.replace("${officeFilePath}", filePath).replace("${outputFilePath}", pdfFilePath.substring(0, dot));
                log.info("word转pdf,cmd={}", cmd);
                Process process = Runtime.getRuntime().exec(cmd);
                // 获取返回状态
                int status = process.waitFor();
                // 销毁process
                process.destroy();
                log.info("word转pdf,file={},pdfFile={},status={}", filePath, pdfFilePath, status);
            } catch (Exception e) {
                log.error("libreoffice word转pdf出错!,file={},{}", filePath, e);
            }
        } else {
            throw new RuntimeException("Unsupport FileType " + filePath);
        }
    }
}
