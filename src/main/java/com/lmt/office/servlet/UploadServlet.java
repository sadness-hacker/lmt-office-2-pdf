package com.lmt.office.servlet;

import com.lmt.office.util.JacobUtils;
import com.lmt.office.util.LibreOfficeUtils;
import lombok.extern.slf4j.Slf4j;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;
import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;

import static com.lmt.office.util.PropertiesUtils.getDownloadUrl;
import static com.lmt.office.util.PropertiesUtils.getPreviewUrl;
import static com.lmt.office.util.PropertiesUtils.getUploadPath;


/**
 *
 * @description 文件上传接口
 * @author bazhandao
 * @date 2019-07-17
 */
@Slf4j
@MultipartConfig
@WebServlet(name = "upload-servlet", value = "/upload")
public class UploadServlet extends HttpServlet {

    /**
     * 是否是windows系统
     */
    private static boolean isWindows = false;

    @Override
    public void init() throws ServletException {
        String os = System.getProperty("os.name");
        isWindows = os != null && os.toUpperCase().contains("windows");
        super.init();
    }

    /**
     * 获取文件上传接口
     * @author bazhandao
     * @date 2019-07-17
     * @param request
     * @param response
     * @throws ServletException
     * @throws IOException
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        request.getRequestDispatcher("/WEB-INF/template/upload.jsp").forward(request, response);
    }

    /**
     * 实现文件上传并转换为pdf功能
     * @author bazhandao
     * @date 2019-07-18
     * @param request
     * @param response
     * @throws ServletException
     * @throws IOException
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            request.setCharacterEncoding("UTF-8");
            String savePath = getUploadPath(request);
            Part part = request.getPart("file");
            String header = part.getHeader("content-disposition");
            String fileName = getFileName(header);
            fileName = newFileName(fileName);
            part.write(savePath + File.separator + fileName);
            String pdfFileName = getPdfFileName(fileName);
            if (isWindows) {
                // Windows系统使用MS Office
                JacobUtils.file2pdf(savePath + File.separator + fileName, savePath + File.separator + pdfFileName);
            } else {
                // 非Windows系统使用LibreOffice
                LibreOfficeUtils.file2pdf(savePath + File.separator + fileName, savePath + File.separator + pdfFileName);
            }
            String dl = getDownloadUrl() + "?file=" + URLEncoder.encode(pdfFileName, "UTF-8");
            String previewUrl = getPreviewUrl() + "?file=" + URLEncoder.encode(pdfFileName, "UTF-8");
            response.setContentType("application/json;charset=UTF-8");
            response.getWriter().write("{\"code\":1,\"msg\":\"success\",\"data\":{\"pdfFileName\":\"" + pdfFileName + "\",\"pdfFileUrl\":\"" + dl + "\",\"previewUrl\":\"" + previewUrl + "\"}}");
        } catch (Exception e) {
            log.error("文件上传转pdf出错!!!", e);
            response.setContentType("application/json;charset=UTF-8");
            response.getWriter().write("{\"code\":9999,\"msg\":\"文件转换pdf出错\"}");
        }
    }

    /**
     * 根据请求头解析出文件名
     * 请求头的格式：火狐和google浏览器下：form-data; name="file"; filename="snmp4j--api.zip"
     * IE浏览器下：form-data; name="file"; filename="E:\snmp4j--api.zip"
     *
     * @param header 请求头
     * @return 文件名
     */
    public String getFileName(String header) {
        /**
         * String[] tempArr1 = header.split(";");代码执行完之后，在不同的浏览器下，tempArr1数组里面的内容稍有区别
         * 火狐或者google浏览器下：tempArr1={form-data,name="file",filename="snmp4j--api.zip"}
         * IE浏览器下：tempArr1={form-data,name="file",filename="E:\snmp4j--api.zip"}
         */
        String[] tempArr1 = header.split(";");
        /**
         * 火狐或者google浏览器下：tempArr2={filename,"snmp4j--api.zip"}
         * IE浏览器下：tempArr2={filename,"E:\snmp4j--api.zip"}
         */
        String[] tempArr2 = tempArr1[2].split("=");
        //获取文件名，兼容各种浏览器的写法
        String fileName = tempArr2[1].substring(tempArr2[1].lastIndexOf("\\") + 1).replaceAll("\"", "");
        fileName = fileName.replaceAll(" ", "");
        return fileName;
    }


    /**
     * 获取pdf文件名
     * @author bazhandao
     * @date 2019-07-17
     * @param fileName
     * @return
     */
    private static String getPdfFileName(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return fileName.substring(0, dot) + ".pdf";
    }

    /**
     * @author bazhandao
     * @date 2019-07-17
     * @param fileName
     * @return
     */
    private static String newFileName(String fileName) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        return sdf.format(new Date()) + "-" + fileName;
    }

}
