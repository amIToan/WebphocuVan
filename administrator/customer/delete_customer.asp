<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 3 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	CMND=Request.Form("CMND")
	Response.Write(CMND)
	if CMND<>"" then
		Dim rs
		set rs=server.CreateObject("ADODB.Recordset")
		sql="delete TaiKhoan where CMND=N'" & CMND & "'"
		rs.open sql,con,1
		sql="delete Account where CMND=N'" & CMND & "'"
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