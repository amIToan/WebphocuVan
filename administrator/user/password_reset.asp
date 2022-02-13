<%session.CodePage=65001%>

<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/md5.asp"-->
<%
	username=Request.QueryString("username")
	key=Request.QueryString("key")
	if key<>md5(username) then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	sql="SELECT UserEmail from [User] where username=N'" & username & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	email=rs("UserEmail")
	rs.close
	
	Userpwd=RandomPassword(20)
	sql="Update [User] set UserPwd='" & md5(Userpwd) & "' where username=N'" & username & "'"
	rs.open sql,con,1
	set rs=nothing
	
	body="Ch&#224;o b&#7841;n " & username & "<br><br>"
	body=body & "B&#7841;n h&#227;y &#273;&#259;ng nh&#7853;p H&#7879; th&#7889;ng v&#7899;i th&#244;ng s&#7889; sau:<br>"
	body=body & "<ul><li>T&#234;n truy nh&#7853;p: " & username & "</li>"
	body=body & "<li>M&#7853;t kh&#7849;u: " & Userpwd & "</li></ul>"
	body=body & "Sau khi &#273;&#259;ng nh&#7853;p b&#7841;n c&#243; th&#7875; &#273;&#7893;i l&#7841;i M&#7853;t kh&#7849;u t&#249;y &#253;.<br><br>"
	sLink="http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("server_port") & "/administrator/"
	body=body & "<a href=""" & sLink & """ target=""_blank""><b>Trang &#273;&#259;ng nh&#7853;p</b></a>"
	
	' Enable UTF-8 -> Unicode translation for form items
	Session.CodePage = 65001 ' UTF-8 code
		
	Set Mail = Server.CreateObject("Persits.MailSender")
	
	Mail.Host = MailServer
		
	Mail.From = AdminMail
   	Mail.FromName = AdminName
   	Mail.AddAddress email
	' message subject
   	Mail.Subject = "Khoi tao lai Mat khau tu http://" & Request.ServerVariables("server_name")
	Mail.Body = body
	Mail.IsHTML = True 
	Mail.CharSet = "UTF-8"
   	Mail.ContentTransferEncoding = "Quoted-Printable"
	Mail.Send 
   	set Mail=nothing
%>
<HTML>
<HEAD>
	<TITLE><%=PAGE_TITLE%></TITLE>
	<META http-equiv=Content-Type content="text/html; charset=utf-8">
</HEAD>
<BODY leftMargin=0 topMargin=0 onLoad="javascript: document.fLogin.uid.focus();">
<%Call header()%>
<div align="center">
  <center>
	<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">
	<tr> 
		<td height="40">&nbsp;</td>
	</tr>
	<tr align="center" valign="middle"> 
		<td align="center" valign="middle">
			<br>
			<font size="4" face="Verdana, Arial, Helvetica, sans-serif">
				<b>Mật khẩu đã được khởi tạo lại</b></font><br><br>
			<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				Bạn hãy kiểm tra mail để lấy mật khẩu mới.<br><br>
				<a href="/administrator/">Trang đăng nhập</a>
			</font> 
		</td>
	</tr>
	<tr> 
		<td height="80">&nbsp;</td>
	</tr>
</table>
  </center>
</div>
<%Call Footer()%>
</BODY>
</HTML>