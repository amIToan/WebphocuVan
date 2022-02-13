<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	id=Request.QueryString("ID")
	sql = "delete Email where ID = "&id
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	set	rs = nothing

	
%>
	<script language="javascript">
		history.back();
		window.opener.location.reload();
	</script>
</body>
</html>
