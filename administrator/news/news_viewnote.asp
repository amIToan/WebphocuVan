<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	NewsId=Request.QueryString("param")
	CatId=Request.QueryString("CatId")
	if not IsNumeric(NewsId) or not IsNumeric(CatId) then
		response.Redirect("/administrator/")
		response.End()
	else
		NewsId=CLng(NewsId)
		CatId=CLng(CatId)
	end if
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ed")
	sql="select Note from News where NewsId=" & NewsId
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		response.End
	end if
	Note=trim(rs("Note"))
	rs.close
	set rs=nothing
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Các 
      ghi chú kèm theo tin</strong></font></td>
  </tr>
  <tr> 
    <td align="left"><div align="justify"><font size="2" face="Arial, Helvetica, sans-serif"><%=Note%></font></div></td>
  </tr>
   <tr> 
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng cửa sổ</a></font></td>
  </tr>
</table>
</body>
</html>
