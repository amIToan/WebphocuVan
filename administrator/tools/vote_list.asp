<%session.CodePage=65001%>
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
	if request.QueryString("CatId")="" or not IsNumeric(request.QueryString("CatId")) then
		CatId=GetFirstCategoryId_With_AP_Role(Session("LstRole"))
	else
		Catid=Clng(request.QueryString("CatId"))
	end if
	
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	if request.QueryString("page")="" or not IsNumeric(request.QueryString("page")) then
		page=1
	else
		page=Clng(request.QueryString("page"))
	end if
	'Field Order
	if request.QueryString("FieldOrder")="" or not IsNumeric(request.QueryString("FieldOrder")) then
		FieldOrder=0 'order by PictureId,CreationDate
		'FieldOrder=1 order by PictureCaption (PictureName)
	else
		FieldOrder=Clng(request.QueryString("FieldOrder"))
	end if

	if request.QueryString("TypeOrder")="" or not IsNumeric(request.QueryString("TypeOrder")) then
		TypeOrder=0 'order by desc
		'TypeOrder=1 order by Ascending
	else
		TypeOrder=Clng(request.QueryString("TypeOrder"))
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Call header()
	Call Menu()
	Title_This_Page="C&#244;ng c&#7909; -> Th&#259;m d&#242; &#253; ki&#7871;n"
	
%>

<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="right" valign="top"> 
    <td height="25"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
      <a href="#">Tìm kiếm</a>&nbsp;| <a href="javascript: winpopup('vote_addnew.asp','<%=CatId%>',420,300);">Thăm 
      dò mới</a> </strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class=normal>
 <tr align="center" bgcolor="FFFFFF"> 
    <td>&nbsp;</td>
	<td align="left"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%=GetNameOfCategory(CatId)%></strong></font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">
		<%sLink=Request.ServerVariables("SCRIPT_NAME") & "?catid=" & catid%>
		Tên thăm dò<a href="<%=sLink%>&FieldOrder=1&TypeOrder=0"><img src="../images/triangle-u.gif" width="16" height="16" border="0" align="absmiddle" title="Z->A"></a><a href="<%=sLink%>&FieldOrder=1&TypeOrder=1"><img src="../images/triangle-d.gif" width="16" height="16" border="0" align="absmiddle"title="A->Z"></a>&nbsp;| 
      	Ngày tạo<a href="<%=sLink%>&FieldOrder=0&TypeOrder=0"><img src="../images/triangle-u.gif" width="16" height="16" border="0" align="absmiddle" title="10 -> 1"></a><a href="<%=sLink%>&FieldOrder=0&TypeOrder=1"><img src="../images/triangle-d.gif" width="16" height="16" border="0" align="absmiddle" title="1 -> 10"></a> 
    </font></td>
  </tr>
  <tr align="center" bgcolor="FFFFFF"> 
    <td valign="top"><%Call ListTreeCategory_WithRole(CatId, "Danh s&#225;ch chuy&#234;n m&#7909;c", "NONE", session("LstRole"), "ap", 0)%></td>
    <td valign="top" colspan="2"><%Call Display_Vote(CatId,Page, FieldOrder,TypeOrder)%></td>
  </tr>
</table>
<%Call Footer()%>
</body>
</html>