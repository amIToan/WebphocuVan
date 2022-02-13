<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<p>
  <%
	Call header()
	
	Title_This_Page="Xóa nhân viên."
	Call Menu()
%>
</p><br><br>
<%
sql = "DELETE FROM Nhanvien WHERE NhanvienID = " & Request.QueryString("NhanvienID")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,3
%>
<br>
<center>
<img src="../../images/icons/brights-brights_icons-delet.gif" align="absmiddle">
Đã xóa cán bộ này!<Br>

<a href="stafflist.asp">Quay về hồ sơ cán bộ</a>
</center>
</body>
</html>
