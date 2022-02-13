<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
Dim rs
IF IsNumeric(request.Form("EventId")) and Cint(request.Form("EventId"))<>0 and Cint(request.Form("CatId"))<>0 and IsNumeric(request.Form("CatId")) THEN
	CatId=Cint(request.Form("CatId"))
	EventId=Cint(request.Form("EventId"))
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	
	'Remove Event from News Table
	set rs=server.CreateObject("ADODB.Recordset")
	sql="delete news where EventId=" & EventId
	rs.open sql,con,1
	'Remove Event from Event Table
	sql="delete Event where EventId=" & EventId
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
		EventId=Cint(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if

	sql="SELECT EventName,CategoryId from Event where EventId=" & EventId
	
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		EventName=rs("EventName")
		CatId=Cint(rs("CategoryId"))
	rs.close

	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	sql="SELECT count(EventId) as dem from News where EventId=" & EventId
	rs.open sql,con,1
		NewsCountRelateWithEvent=Cint(rs("dem"))
	rs.close
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
    	<%=EventName%>
    </strong></font></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="40" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif">
		<%if NewsCountRelateWithEvent<>0 then
			Response.Write "<br>Có <strong>" & NewsCountRelateWithEvent & "</strong> tin thuộc sự kiện.<br>"
			Response.Write "Đồng thời xóa luôn sự kiện khỏi các tin trên?"
		else
			Response.Write "<br>Bạn chắc chắn muốn xóa sự kiện này?"
		end if%>
	</font> </td>
  </tr>
  <tr> 
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
		<a href="javascript: window.document.fDelete.submit();">Xóa sự kiện</a> </font></td>
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng 
      cửa sổ</a></font></td>
  </tr>
</table>
<input type="hidden" name="EventId" value="<%=EventId%>">
<input type="hidden" name="CatId" value="<%=CatId%>">
</form>
</body>
</html>
