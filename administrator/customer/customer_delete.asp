<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	CMND= trim(Request.QueryString("param"))
%>
<html>
<head>
<title>Xóa</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fDelete" method="post" action="delete_customer.asp?del=1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center" valign="middle"> 
    <td height="30"><font size="3" face="Arial, Helvetica, sans-serif">
		<strong><br>
		<img src="../../images/icons/System-Recycle-Bin-Empty-icon.jpg" width="48" height="48" align="absmiddle">Chắc chắn xóa ?<br></strong>
		&#8226;&nbsp;<%= CMND %>
	</font></td>
  </tr>
  <tr align="center" valign="middle">
    <td height="30"><br><font size="2" face="Arial, Helvetica, sans-serif">
		<a href="javascript: document.fDelete.submit();">Chắc chắn</a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="javascript: window.close();">Đóng cửa sổ</a>
	</font>
	<input type="hidden" name="CMND" value="<%=CMND%>" >
	</td>
  </tr>
</table>
</form>
</body>
</html>
