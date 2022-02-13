<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_human")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	lang=Request.QueryString("lang")
	lang=replace(lang,"'","''")
	action =trim(Request.Querystring("action"))
	action=replace(action,"'","''")
	if action = "" or action=Null then
		action =trim(Request.Form("action"))
		action=replace(action,"'","''")
	end if
%>
<html>
<head>

<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<p>
  <%
select case action
  case "update"
  	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT Count(ChucDanhID) as iCount FROM ChucDanh "
	rs.open sql,con,1
	i =	rs("iCount")
	j=1
	for j= 1 to i
		IDChucDanh=Trim(Request.Form("IDChucDanh"&j))
		DesChucDanh=Trim(Request.Form("txtChucDanh"&j))
		sqlCD 	=	"update ChucDanh set Description=N'"& DesChucDanh &"'"
		sqlCD = sqlCD + "where ChucDanhID = '"&IDChucDanh&"'"
		Set rsCD = Server.CreateObject("ADODB.Recordset")
		rsCD.open sqlCD,con,1
		set	rsCD = nothing
	next
	i=i+1
	IDChucDanh=Trim(Request.Form("IDChucDanh"&i))
	DesChucDanh=Trim(Request.Form("txtChucDanh"&i))
	if  (DesChucDanh <>"" or DesChucDanh <> NULL) and (IDChucDanh<>0 or IDChucDanh<>"") then
		sqlCD =	"insert into ChucDanh(Description) values(N'"& DesChucDanh &"')"
		Set rsCD = Server.CreateObject("ADODB.Recordset")
		rsCD.open sqlCD,con,1
		set	rsCD = nothing
	end if
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_	
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
	%>
	<script language="javascript">
		history.back();
	</script>
	<%
 case "del"
 	IDChucDanh=Trim(Request.QueryString("IDChucDanh"&j))
	sqlCD="delete ChucDanh where ChucDanhID='" & IDChucDanh & "'"
	Set rsCD = Server.CreateObject("ADODB.Recordset")
	rsCD.open sqlCD,con,1
	set	rsCD = nothing
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
end select

%>
</p>
</body>
</html>
