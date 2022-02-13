<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>	
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/md5.asp"-->
<%
	oldpwd=request.Form("oldpwd")
	newpwd=request.Form("newpwd")
	newpwdcon=request.Form("newpwdcon")
	booError=false
	sError=""
	Dim rs
	set rs=Server.CreateObject("ADODB.Recordset")
	
	sql="SELECT count(*) as dem from [USER] where username=N'" & session("user") & "' and Userpwd='" & md5(oldpwd) & "'"
	rs.open sql,con,1
	if Clng(rs("dem"))=0 then
		sError=sError & "&nbsp;-&nbsp;M&#7853;t kh&#7849;u c&#361; kh&#244;ng ch&#237;nh x&#225;c<br>"
		booError=true
	end if
	rs.close
	
	If newpwd<>newpwdcon then
		sError=sError & "&nbsp;-&nbsp;M&#7853;t kh&#7849;u g&#245; l&#7841;i kh&#244;ng ch&#237;nh x&#225;c<br>"
		booError=True
	end if
	
	If len(newpwd)<6 then
		sError=sError & "&nbsp;-&nbsp;M&#7853;t kh&#7849;u ph&#7843;i nhi&#7873;u h&#417;n 6 k&#253; t&#7921;<br>"
		booError=True
	end if

	if not booError then
		sql="Update [USER] set Userpwd='" & md5(newpwd) & "' where username=N'" & session("user") & "'"
		rs.open sql,con,1
	end if
	
	set rs=nothing
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript"><!--
	function clearpwd()
	{
		window.opener.document.fChangePwd.oldpwd.value="";
		window.opener.document.fChangePwd.newpwd.value="";
		window.opener.document.fChangePwd.newpwdcon.value="";
		self.close();
	}
//--></script>
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
<%else%>
  <tr> 
    <td colspan="2" align="center"><font size="3" face="Arial, Helvetica, sans-serif"><strong>
		<br>Đổi mật khẩu thành công!
	</strong></font></td>
  </tr>
<%End if%>
  <tr align="center"> 
    <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: clearpwd();"><br>
      Đóng cửa sổ</a></font></td>
  </tr>
</table>

</body>
</html>
