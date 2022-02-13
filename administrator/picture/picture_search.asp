<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_picture.asp" -->
<%
f_permission = administrator(false,session("user"),"m_editor")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	=	"../../images/icons/iPhoto-icon.jpg"
	Title_This_Page="Quản lý ảnh."
	Call header()
	Call Menu()	
	Call Picture_list()
%>

<%Call Footer()%>
</body>
</html>
