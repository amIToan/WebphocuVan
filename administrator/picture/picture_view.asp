<%session.CodePage=65001%>
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/constant.asp"-->
<html>
<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onLoad="window.resizeTo(document.images[0].width+30,document.images[0].height+ 35)">
<%param=replace(request.QueryString("param"),"@","\")%>
<center><a href="javascript: window.close();">
	<%if Instr(param,":\")>0 then%>
		<img src="<%=param%>" border="0">
	<%else%>
		<img src="<%=NewsImagePath%><%=param%>" border="0">
	<%end if%>
</a></center>
</body></html>