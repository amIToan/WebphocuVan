<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
txtGuide=""
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	addOrEddit=	GetNumeric(Request.QueryString("addOrEddit"),0)
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%
	Title_This_Page="Khách hàng -> Cập nhật email khách hàng"
	img ="../../images/icons/icon_customer1.gif"
	Call header()
	
%>
<form name="fEmail" method="post" action="upListEmail.asp" target="_blank">
<table width="900" border="0" align="center" cellpadding="1" cellspacing="1" style="border:#CCCCCC solid 1px;">

<%
namek 		= 	""
NgaySinh = now()
IDNhom		=	0
XungHo		=	""
DiaChi		=	""
Tel			=	""
Email		=	""
GhiChu		=	""
if addOrEddit = 1 then
	ID = Request.QueryString("ID")
	sql = "SELECT * FROM Email where ID = "& ID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	If not rs.eof Then
		namek 		= 	Trim(rs("Ten"))
		IDTamLy     =   rs("IDTamLy")	
		IDCongViec	=   rs("IDCongViec")
		NgaySinh	=	rs("NgaySinh")
		if isDate(NgaySinh)  = false  then
			NgaySinh = now()
		end if
		IDXungHo	=	rs("IDXungho")
		DiaChi		=	Trim(rs("Diachi"))
		Tel			=	Trim(rs("DienThoai"))
		Email		=	trim(rs("Email"))
		iDis		=	rs("Disabled")
		GhiChu		=	Trim(rs("Ghichu"))
	End If
	set rs = nothing
end if
%>
	<tr>
		<td width="900"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
          <tr>
            <td class="CTieuDe" align="center"><p>		
			
			    UPDATE LIST EMAIL</p>            </td>
          </tr>
          <tr>
            <td class="CTieuDe">
                Job:
              <%call fshowCustomGroup(IDCongViec,"sel_cong_viec",2)%>
            </td>
          </tr>
          <tr>
            <td valign="middle">
                <br />
                <div class="Ctieudenho">Note: Copy email list </div>
                <div class="CSubTitle">when you copy the email list then you use by WordPad which is not Word-Excel <br />please input each email one line</div>
<br />
              <textarea name="txtGhiChu" cols="60" rows="30" class="CTxtContent" id="txtGhiChu"><%=GhiChu%></textarea></td>
          </tr>
            <tr>
                <td align="center">
                    			<input type="hidden" name="iLenCSKH" value="<%=sTT-1%>">
			<input name="ID" type="hidden" value="<%=ID%>">
			<input name="addOrEddit" type="hidden" value="<%=addOrEddit%>">
			<input type="submit" name="Submit22" value=" Update " >
                </td>
            </tr>
</td>
	</tr>
</table></form>
</body>
</html>
