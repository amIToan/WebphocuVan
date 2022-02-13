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
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
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
	sql="SELECT Count(PhongID) as iCount FROM PhongBan "
	rs.open sql,con,1
	i =	rs("iCount")
	j=1
	for j= 1 to i
		IDPhongBan=Trim(Request.Form("IDPhongBan"&j))
		DesPhongBan=Trim(Request.Form("txtPhongBan"&j))
		sqlCD 	=	"update PhongBan set Description=N'"& DesPhongBan &"'"
		sqlCD = sqlCD + " where PhongID = '"&IDPhongBan&"'"
		Set rsCD = Server.CreateObject("ADODB.Recordset")
		Response.Write(sqlCD&"<br>")
		rsCD.open sqlCD,con,1
		set	rsCD = nothing
	next
	i=i+1
	
	IDPhongBan=Trim(Request.Form("IDPhongBan"&i))
	DesPhongBan=Trim(Request.Form("txtPhongBan"&i))
	if  (DesPhongBan <>"" or DesPhongBan <> NULL) and IDPhongBan<>"" then
		sqlCD =	"insert into PhongBan(PhongID,Description) values('"&IDPhongBan&"',N'"& DesPhongBan &"')"
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
 	IDPhongBan=Trim(Request.QueryString("IDPhongBan"&j))
	sqlCD="delete PhongBan where PhongID='" & IDPhongBan & "'"
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
