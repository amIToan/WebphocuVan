<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_editor")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
Dim rs
IF IsNumeric(request.Form("PicId")) and CLng(request.Form("PicId"))<>0 and CLng(request.Form("CatId"))<>0 and IsNumeric(request.Form("CatId")) THEN
	CatId=CLng(request.Form("CatId"))
	PicId=CLng(request.Form("PicId"))
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	
	'Remove Picture from News Table
	set rs=server.CreateObject("ADODB.Recordset")
	sql="delete news where PictureId=" & PicId
	rs.open sql,con,1
	'Remove Picture from Picture Table
	sql="delete picture where PictureId=" & PicId
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
		PicId=CLng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if

	sql="SELECT SmallPictureFilename,CategoryId from Picture where PictureId=" & PicId
	
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		imagespath=rs("SmallPictureFilename")
		CatId=CLng(rs("CategoryId"))
	rs.close

	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	sql="SELECT count(PictureId) as dem from News where PictureId=" & PicId
	rs.open sql,con,1
		NewsCountRelateWithPic=CLng(rs("dem"))
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
    <td colspan="2"><img src="<%=NewsImagePath%><%=imagespath%>"></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="40" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif">
		<%if NewsCountRelateWithPic<>0 then
			Response.Write "<br>Có <strong>" & NewsCountRelateWithPic & "</strong> tin lấy ảnh trên làm minh họa.<br>"
			Response.Write "Đồng thời xóa luôn ảnh khỏi các tin trên?"
		else
			Response.Write "<br><strong>Bạn chắc chắn muốn xóa ảnh này?</strong>"
		end if%>
	</font> </td>
  </tr>
  <tr> 
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
		<a href="javascript: window.document.fDelete.submit();">Xóa ảnh</a> </font></td>
    <td width="50%" height="25" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng 
      cửa sổ</a></font></td>
  </tr>
</table>
<input type="hidden" name="PicId" value="<%=PicId%>">
<input type="hidden" name="CatId" value="<%=CatId%>">
</form>
</body>
</html>
