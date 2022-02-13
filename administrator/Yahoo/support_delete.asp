<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	idNhanvien=Request.QueryString("id")
	set rs = server.CreateObject("ADODB.Recordset")
	if not IsNumeric(idNhanvien) then
		response.Redirect("/administrator/")
		response.End()
	else
		idNhanvien=CLng(idNhanvien)
	end if
	if request.Form("action")="Delete" then
		sql=" delete from SupportYahoo where id ="& idNhanvien
		rs.open sql,con,1
		
		set rs=nothing
		response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
		"window.opener.location.reload();" & vbNewline &_
		"window.close();" & vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline
		response.End()
	else
		sql="select * from SupportYahoo where id ="& idNhanvien
		rs.open sql,con,1
		if rs.eof then
			rs.close
			set rs=nothing
			response.End()
		end if
		Hoten=rs("Hoten")
		Chucvu=rs("ghichu")
		rs.close
		set rs=nothing
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fDelete" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>" method="post">
<table width="98%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td height="25" colspan="2" align="center" valign="middle"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Bạn 
        chắc chắn muốn xóa nhân viên: </strong></font></p></td>
  </tr>
  <tr> 
    <td colspan="2" align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=Hoten%></font></td>
  </tr>
  <tr> 
    <td colspan="2" align="center"><font size="2" face="Arial, Helvetica, sans-serif"><b>Chức vụ:&nbsp;</b><%=Chucvu%></font></td>
  </tr>
  <tr> 
    <td width="50%" align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a class="w3-btn w3-red w3-round" href="javascript: document.fDelete.submit();"><i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Xóa 
      tin</a></font></td>
    <td width="50%" align="center"><a class="w3-btn w3-blue w3-round" href="javascript: window.close();"><font size="2" face="Arial, Helvetica, sans-serif"><i class="fa fa-times" aria-hidden="true"></i>
 Đóng 
      cửa sổ</font></a></td>
  </tr>
</table>
<input type="hidden" name="action" value="Delete">
</form>
</body>
</html>
