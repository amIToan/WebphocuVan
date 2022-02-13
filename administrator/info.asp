<%session.CodePage=65001%>
<%
	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<HTML>
	<HEAD>
		<TITLE><%=PAGE_TITLE%></TITLE>
		<META http-equiv=Content-Type content="text/html; charset=utf-8">
	</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	=	"../images/icons/delete-48x48.gif"
	Title_This_Page="Chuc nang chua duoc cap"
	Call header()
	Call Menu()
%>
<%Call Footer()%>
</BODY>
</HTML>