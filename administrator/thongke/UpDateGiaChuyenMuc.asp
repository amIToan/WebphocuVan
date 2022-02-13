<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%	
	
	stt			=	GetNumeric(Request.Form("stt"),0)
	for i = 1 to stt 
		NewsID	=	GetNumeric(Request.Form("NewsID"&i),0)
		Gia = 	Chuan_money(Request.Form("Gia"&i))
		set rs=server.CreateObject("ADODB.Recordset")
		sql = "Update News set"
		sql	=	sql	+	" Gia	='"& Gia &"'"
		sql	=	sql	+	" Where NewsID="& NewsID	
		Response.Write(i&sql&"<br>")		
		rs.open sql,con,3
		set rs=nothing	
	
	next		
%>		
<script language="JavaScript">
	alert("Đã cập nhật")
	window.close();
</script>	
