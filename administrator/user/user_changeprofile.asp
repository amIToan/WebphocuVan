<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>	
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
	Email=request.Form("Email")
	FullName=request.Form("FullName")
	Title=request.Form("Title")
	booError=false
	sError=""
	
	If len(Email)<6 then
		sError=sError & "&nbsp;-&nbsp;Email không chính xác<br>"
		booError=True
	end if

	if not booError then
		Dim rs
		set rs=Server.CreateObject("ADODB.Recordset")
		sql="Update [USER] set "
		sql=sql & "UserEmail='" & Email & "'"
		sql=sql & ",UserFullName=N'" & FullName & "'"
		sql=sql & ",UserTitle=N'" & Title & "'"
		sql=sql & " WHERE username=N'" & session("user") & "'"
		
		rs.open sql,con,1
		set rs=nothing
	end if
	
	
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="96%" border="0" cellspacing="2" cellpadding="2" align="center">
<%if booError then%>
  <tr> 
    <td colspan="2"><font size="3" face="Arial, Helvetica, sans-serif"><strong><font color="#FF0000">&nbsp;&nbsp;Có 
      lỗi:</font></strong></font></td>
  </tr>
  <tr> 
    <td width="10%">&nbsp;</td>
    <td width="90%"><font size="2" face="Arial, Helvetica, sans-serif"><%=sError%></font></td>
  </tr>
  <tr align="center"> 
    <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();"><br>
      Đóng cửa sổ</a></font></td>
  </tr>
<%else%>
  <tr> 
    <td colspan="2" align="center"><font size="3" face="Arial, Helvetica, sans-serif"><strong>
		<br>Sửa thông tin cá nhân thành công!
	</strong></font></td>
  </tr>
  <tr align="center"> 
    <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.opener.location.reload();window.close();"><br>
      Đóng cửa sổ</a></font></td>
  </tr>
<%End if%>
</table>

</body>
</html>
