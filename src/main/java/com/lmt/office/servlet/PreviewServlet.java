package com.lmt.office.servlet;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * @description 文件预览Servlet
 * @author bazhandao
 * @date 2019-07-17
 */
@WebServlet(name = "preview-servlet", value = "/preview")
public class PreviewServlet extends HttpServlet {


    /**
     * pdf文件预览接口
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
        String fileName = request.getParameter("file");
        request.setAttribute("fileName", fileName);
        request.getRequestDispatcher("/WEB-INF/template/preview.jsp").forward(request, response);
    }
}
