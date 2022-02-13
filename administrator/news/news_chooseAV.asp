<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_av.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<br>
<%Call AudioVideo_List("DESC")%>
<p align="center"><font face="Arial, Helvetica, sans-serif" size="2"><a href="javascript: window.close();">Đóng cửa sổ</a></font></p>
</body>
</html>
