<%-- 
    Document   : Jexcel2
    Created on : 2017-6-15, 21:12:44
    Author     : nifflers
--%>

<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@page import="java.io.*"%>
<%@page import="net.nifflers.jexcel.*"%>
<!DOCTYPE html>
<%
    String name = "TestExcel";
    OutputStream os = response.getOutputStream();
    response.reset(); //清空输出流。

    //对中文文件名称处理
    response.setCharacterEncoding("UTF-8");//设置相应内容的编码格式
    name = java.net.URLEncoder.encode(name, "UTF-8");
    response.setHeader("Content-Disposition", "attachment;filename=" + new String(name.getBytes("UTF-8"), "GBK") + ".xls");
    response.setContentType("application/msexcel");//定义输出类型
    
    Jexcel2 excel2 = new Jexcel2();
    excel2.createExcel(os);
    
%>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>JSP Page</title>
    </head>
    <body>
        <h1>Hello World!</h1>
    </body>
</html>
