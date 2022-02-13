 <%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_order_output")
if f_permission < 3 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
varEmail     	= Trim(Request.Form("txt_mail"))
txt_form		= Trim(Request.Form("txt_form"))
strCC			= Trim(Request.Form("txtCC"))
strbcc			= Trim(Request.Form("txtBcc"))	
txt_subject  	= Trim(Request.Form("txt_subject"))
txt_content  	= Trim(Request.Form("txt_content"))

 if varEmail<>"" and txt_subject<>"" and txt_content<>"" then 
  	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = MailServer
	Mail.From = txt_form ' From address
	Mail.FromName = "Sach xbook.com.vn" ' optional
	Mail.AddAddress varEmail
	Mail.Subject = txt_subject
	Mail.Body = txt_content
	Mail.IsHTML = True 
	Mail.CharSet = "UTF-8"
	Mail.ContentTransferEncoding = "Quoted-Printable"
	Mail.Send
  
   response.Write "<script language=""JavaScript"">" & vbNewline &_
				"<!--" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline 
	Response.Write("Da gui thanh cong email")
	Response.Write("<div align=""center""><a href=""javascript:history.back()"">Back</a></div>&nbsp;&nbsp;") 	
 else
		Response.Write("Noi dung gui email khong hop le!")
		Response.Write("<div align=""center""><a href=""javascript:history.back()"">Back</a></div>&nbsp;&nbsp;") 	
 end if
%>

  
   