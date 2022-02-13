<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
	
ID = Request.QueryString("param") 
sqlCD="Delete ColorTable where ID = '"&ID&"'"
Set rsCD = Server.CreateObject("ADODB.Recordset")
rsCD.open sqlCD,con,1
set	rsCD = nothing

Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists(server.MapPath(Request.QueryString("PathColor"))) then
  fs.DeleteFile(server.MapPath(Request.QueryString("PathColor")))
end if
set fs=nothing
		
Response.Write	"<script language=""JavaScript"">" & vbNewline &_
		"	<!--" & vbNewline &_
		"		alert('Đã cập nhật thành công');" & vbNewline &_
		"		window.opener.location.reload();" & vbNewline &_
		"		window.close();" & vbNewline &_	
		"	//-->" & vbNewline &_
		"</script>" & vbNewline	
%>