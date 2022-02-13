<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission= administrator(false,session("user"),"m_user")
if f_permission <3 then
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
		"	<!--" & vbNewline &_
		"		alert('Bạn không có quyền xóa!');" & vbNewline &_
		"		window.close();" & vbNewline &_
		"	//-->" & vbNewline &_
		"</script>" & vbNewline
	response.End()
end if
%>
<%
	UserName=Request.QueryString("param")
	UserName=Replace(UserName,"'","''")
	
	if request.Form("confirm")="Yes" and UserName<>"" then
		Dim rs
		set rs=server.CreateObject("ADODB.Recordset")
		sql="delete UserDistribution where Username=N'" & username & "'"
		rs.open sql,con,1
		sql="delete [User] where Username=N'" & username & "'"
		rs.open sql,con,1
		set rs=nothing
		
		Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
		response.End()
	end if
%>
<html>
<head>
<title>Chuyen muc moi</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fDelete" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center" valign="middle"> 
    <td height="30"><font size="3" face="Arial, Helvetica, sans-serif">
		<strong><br>Chắc chắn xóa User?<br></strong>
		&#8226;&nbsp;<%=Username%>
	</font></td>
  </tr>
  <tr align="center" valign="middle">
    <td height="30"><br><font size="2" face="Arial, Helvetica, sans-serif">
		<a href="javascript: document.fDelete.submit();">Chắc chắn</a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="javascript: window.close();">Đóng cửa sổ</a>
	</font></td>
  </tr>
</table>
<input type="hidden" name="confirm" value="Yes">
</form>
</body>
</html>
