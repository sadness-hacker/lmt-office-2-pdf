package com.lmt.office.servlet;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

import static com.lmt.office.util.PropertiesUtils.getUploadPath;


/**
 * @description 文件下载Servlet
 * @author bazhandao
 * @date 2019-07-17
 */
@WebServlet(name = "download-servlet", value = "/download")
public class DownloadServlet extends HttpServlet {

    /**
     * 实现文件下载接口
     * @author bazhandao
     * @date 2019-07-17
     * @param request
     * @param response
     * @throws ServletException
     * @throws IOException
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        request.setCharacterEncoding("UTF-8");
        String savePath = getUploadPath(request);
        String fileName = request.getParameter("file");
        if (fileName.toLowerCase().endsWith(".pdf")) {
            response.setContentType("application/pdf");
            response.addHeader("Content-Disposition", "filename=" + new String(fileName.getBytes(), StandardCharsets.ISO_8859_1));
        } else {
            response.setContentType("multipart/form-data");
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), StandardCharsets.ISO_8859_1));
        }
        byte [] bytes = Files.readAllBytes(Paths.get(savePath + File.separator + fileName));
        response.addHeader("Content-Length", "" + bytes.length);
        response.getOutputStream().write(bytes);
    }
}
