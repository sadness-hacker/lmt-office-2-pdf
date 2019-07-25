<%@ page contentType="text/html;charset=UTF-8"%>

<%
    String os = System.getProperty("os.name");
    response.getWriter().print(os);
%>
