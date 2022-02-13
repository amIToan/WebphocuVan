<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	if request.QueryString("CatId")="" or not IsNumeric(request.QueryString("CatId")) then
		CatId=1
	else
		Catid=Cint(request.QueryString("CatId"))
	end if
	if request.QueryString("page")="" or not IsNumeric(request.QueryString("page")) then
		page=1
	else
		page=Cint(request.QueryString("page"))
	end if
	'Field Order
	if request.QueryString("FieldOrder")="" or not IsNumeric(request.QueryString("FieldOrder")) then
		FieldOrder=0 'order by PictureId,CreationDate
		'FieldOrder=1 order by PictureCaption (PictureName)
	else
		FieldOrder=Cint(request.QueryString("FieldOrder"))
	end if

	if request.QueryString("TypeOrder")="" or not IsNumeric(request.QueryString("TypeOrder")) then
		TypeOrder=0 'order by desc
		'TypeOrder=1 order by Ascending
	else
		TypeOrder=Cint(request.QueryString("TypeOrder"))
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="right" valign="top"> 
    <td height="25" align="left" valign="middle"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&nbsp;&nbsp; 
      <a href="javascript: winpopup('/administrator/picture/picture_addnew.asp','<%=CatId%>',420,300);">Tạo ảnh mới</a></strong></font></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class=normal>
  <tr align="center" bgcolor="FFFFFF"> 
    <td>&nbsp;</td>
	<td align="left"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%=GetNameOfCategory(CatId)%></strong></font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">
		<%sLink=Request.ServerVariables("SCRIPT_NAME") & "?catid=" & catid%>
		Tên ảnh<a href="<%=sLink%>&FieldOrder=1&TypeOrder=0"><img src="../images/triangle-u.gif" width="16" height="16" border="0" align="absmiddle" title="Z->A"></a><a href="<%=sLink%>&FieldOrder=1&TypeOrder=1"><img src="../images/triangle-d.gif" width="16" height="16" border="0" align="absmiddle"title="A->Z"></a>&nbsp;| 
      	Ngày cập nhật<a href="<%=sLink%>&FieldOrder=0&TypeOrder=0"><img src="../images/triangle-u.gif" width="16" height="16" border="0" align="absmiddle" title="10 -> 1"></a><a href="<%=sLink%>&FieldOrder=0&TypeOrder=1"><img src="../images/triangle-d.gif" width="16" height="16" border="0" align="absmiddle" title="1 -> 10"></a> 
    </font></td>
  </tr>
  <tr align="center" bgcolor="FFFFFF"> 
    <td valign="top"><%Call ListTreeCategory(CatId)%></td>
    <td valign="top" colspan="2"><%Call Display_Images_Library(CatId,1, FieldOrder,TypeOrder,1)%></td>
  </tr>
</table>
<p align="center"><font face="Arial, Helvetica, sans-serif" size="2"><a href="javascript: window.close();">Đóng cửa sổ</a></font></p>
</body>
</html>
