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
Call AuthenticateWithRole(AudioVideoCategoryId,Session("LstRole"),"ap")
	
Dim rs
IF IsNumeric(request.Form("Av_id")) and Clng(request.Form("Av_id"))<>0  THEN
	Av_id=Clng(request.Form("Av_id"))

	set rs=server.createObject("ADODB.Recordset")
	sql="delete AudioVideo where Av_id=" & Av_id
	rs.open sql,con,1

	set rs=nothing
	response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.opener.focus();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	response.End()
ELSE
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		Av_id=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	
	sql="SELECT	Av_id, Av_Title " &_
		"FROM	AudioVideo " &_
		"WHERE     (Av_id = " & Av_id & ")"
		
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if not rs.eof then
		Av_Title=rs("Av_Title")
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	rs.close
	set rs=nothing
END IF
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fDelete">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td colspan="2" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
    	<strong><%=Av_Title%></strong>
    </font></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="40" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif">
		Bạn chắc chắn muốn xóa Videoclips này?
	</font> </td>
  </tr>
  <tr> 
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
		<a class="w3-btn w3-red w3-round" href="javascript: window.document.fDelete.submit();">
            <i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Xóa VideoClips</a> </font></td>
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
        <a class="w3-btn w3-blue w3-round" href="javascript: window.close();">
            <i class="fa fa-times" aria-hidden="true"></i> Ðóng cửa sổ</a></font></td>
  </tr>
</table>
<input type="hidden" name="Av_id" value="<%=Av_id%>">
</form>
</body>
</html>
