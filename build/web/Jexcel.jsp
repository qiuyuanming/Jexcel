<%-- 
    Document   : test
    Created on : 2017-6-15, 21:09:20
    Author     : nifflers
--%>

<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@page import="java.io.*" %>
<%@page import="net.nifflers.jexcel.*"%>
<!DOCTYPE html>
<% 
    String fname = "学校竞争力情况";
    OutputStream os = response.getOutputStream();//取得输出流
    response.reset();//清空输出流
    
    //下面是对中文文件名的处理
    response.setCharacterEncoding("UTF-8");//设置相应内容的编码格式
    fname = java.net.URLEncoder.encode(fname,"UTF-8");
    response.setHeader("Content-Disposition","attachment;filename="+new String(fname.getBytes("UTF-8"),"GBK")+".xls");
    response.setContentType("application/msexcel");//定义输出类型
    Jexcel sw = new Jexcel();
    sw.createExcel(os);

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
