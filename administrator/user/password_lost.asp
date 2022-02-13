<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp"-->
<!--#include virtual="/administrator/inc/md5.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
	if Trim(request.form("uid"))<>"" and Trim(request.form("email"))<>"" then
		username=Trim(request.form("uid"))
		email=Trim(request.form("email"))
		Dim rs
		set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from [User] where UserName=N'" & UserName & "' and UserEmail='" & email & "'"
		rs.open sql,con,1
		if rs.eof then
			result="None"
		else
			body="Ch&#224;o b&#7841;n " & username & "<br><br>"
			body=body & "V&#236; b&#7841;n &#273;&#227; kh&#244;ng nh&#7899; &#273;&#432;&#7907;c m&#7853;t kh&#7849;u c&#7911;a m&#236;nh v&#224; y&#234;u c&#7847;u H&#7879; th&#7889;ng l&#7845;y l&#7841;i m&#7853;t kh&#7849;u, n&#234;n b&#7841;n nh&#7853;n &#273;&#432;&#7907;c Email n&#224;y.<br><br>"
			body=body & "H&#227;y nh&#7845;n v&#224;o &#273;&#432;&#7901;ng Link d&#432;&#7899;i &#273;&#226;y &#273;&#7875; ti&#7871;p t&#7909;c qu&#225; tr&#236;nh l&#7845;y l&#7841;i m&#7853;t kh&#7849;u:<br><br>"
			sLink=Replace(Request.ServerVariables("SCRIPT_NAME"),"password_lost.asp","password_reset.asp")
			sLink="http://" & Request.ServerVariables("server_name") & ":" & Request.ServerVariables("server_port") & sLink & "?username=" & username & "&key=" & md5(username)
			body=body & "<a href=""" & sLink & """ target=""_blank"">" & sLink & "</a>"

			' Enable UTF-8 -> Unicode translation for form items
			Session.CodePage = 65001 ' UTF-8 code
		
			Set Mail = Server.CreateObject("Persits.MailSender")
   			' enter valid SMTP host
   			Mail.Host = MailServer
		
	   		Mail.From = AdminMail
   			Mail.FromName = AdminName
   			Mail.AddAddress email
	   		' message subject
   			Mail.Subject = "Lay lai Mat khau tu http://" & Request.ServerVariables("server_name")
	   		Mail.Body = body
			Mail.IsHTML = True 
	   		Mail.CharSet = "UTF-8"
   			Mail.ContentTransferEncoding = "Quoted-Printable"
	   		Mail.Send 
   			set Mail=nothing
   			Mailsend=true
			result="CheckMail"
		end if
	End if
%>
<HTML>
<HEAD>
	<TITLE><%=PAGE_TITLE%></TITLE>
	<META http-equiv=Content-Type content="text/html; charset=utf-8">
</HEAD>
<BODY leftMargin=0 topMargin=0>
<%Call header()%>
<div align="center">
  <center>
	<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fSend">
	  <table border="0" cellspacing="2" cellpadding="2" align="center">
        <tr> 
          <td height="60">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <%if result="None" then%>
        <tr> 
          <td height="30" valign="top" colspan="2"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000"><strong>* 
            Không tìm thấy tài khoản</strong></font> </td>
        </tr>
        <%End if%>
        <%if result="CheckMail" then%>
        <tr> 
          <td colspan="2" align="center"> <br> <font size="4" face="Verdana, Arial, Helvetica, sans-serif"> 
            <b>Bạn hãy làm theo hướng dẫn</b></font><br> <br> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
            được gửi vào địa chỉ <b><%=Email%></b><br>
            để tiếp tục quá trình lấy lại mật khẩu.<br><br>
			<a href="/administrator/">Trang đăng nhập</a></font> </td>
        </tr>
        <%Else%>
        <tr> 
          <td colspan="2" bgcolor="#416aa9"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>QUÊN 
            MẬT KHẨU</strong></font></td>
        </tr>
        <tr> 
          <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
            truy nhập:</font></td>
          <td><input name="uid" type="text" id="uid" size="25"></td>
        </tr>
        <tr> 
          <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Email:</font></td>
          <td><input name="email" type="text" id="email" size="25"></td>
        </tr>
        <tr> 
          <td colspan="2" align="center"><input type="submit" name="Submit" value="Gửi yêu cầu">
            <input type="button" name="Submit2" value="Trang đăng nhập" onclick="javascript: window.open('/administrator/','_self');"></td>
        </tr>
        <%End if%>
        <tr> 
          <td height="80">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
	</form>
  </center>
</div>
<%Call Footer()%>
</BODY>
</HTML>