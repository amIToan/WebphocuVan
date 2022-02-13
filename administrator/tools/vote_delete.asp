<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
Dim rs
IF IsNumeric(request.Form("VoteId")) and Clng(request.Form("VoteId"))<>0 and Clng(request.Form("CatId"))<>0 and IsNumeric(request.Form("CatId")) THEN
	CatId=Clng(request.Form("CatId"))
	VoteId=Clng(request.Form("VoteId"))
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	
	'Remove VoteItem from VoteItem Table
	set rs=server.CreateObject("ADODB.Recordset")
	sql="delete VoteItem where VoteId=" & VoteId
	rs.open sql,con,1
	'Remove Vote from Vote Table
	sql="delete Vote where VoteId=" & VoteId
	rs.open sql,con,1
	set rs=nothing
	response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	response.End()
ELSE
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		VoteId=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if

	sql="SELECT VoteTitle,CategoryId from Vote where VoteId=" & VoteId
	
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		VoteTitle=rs("VoteTitle")
		CatId=Clng(rs("CategoryId"))
	rs.close

	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	set rs=nothing
END IF
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fDelete">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td colspan="2" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><strong>
    	<%=VoteTitle%>
    </strong></font></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="40" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif">
		Bạn chắc chắn muốn xóa thăm dò này?
	</font> </td>
  </tr>
  <tr> 
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
		<a href="javascript: window.document.fDelete.submit();">Xóa thăm dò</a> </font></td>
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Ðóng 
      cửa sổ</a></font></td>
  </tr>
</table>
<input type="hidden" name="VoteId" value="<%=VoteId%>">
<input type="hidden" name="CatId" value="<%=CatId%>">
</form>
</body>
</html>
