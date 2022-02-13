<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	vnd =chuan_money(Request.Form("vnd"))
	sql = "update exchange set VND = "&vnd
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1


%>
<script language="javascript">
	window.close();
	window.opener.location.reload();
</script>
</body>

</html>