<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/Home-icon.jpg"
	Title_This_Page="Danh sách sách tiêu biểu -> Liệt kê"
	Call header()
	Call Menu()
	
	
%>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="bottom" class="CTieuDeNhoNho" align="center"><img src="../../images/icons/Task-List-40x40.gif" width="40" height="40" align="absmiddle"> &nbsp;TIN TỨC &amp; SẢN PHẨM TIÊU BIỂU THEO CHỦ ĐỀ <img src="../../images/icons/Task-List-40x40.gif" width="40" height="40" align="absmiddle"></td>
  </tr>
  <tr>
    <td >
	<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
      <tr>
        <td width="4%" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><strong>STT</strong></td>
        <td width="53%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tiêu đề </strong></td>
        <td width="4%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>View </strong></td>
        <td width="8%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong><a href="../KhuyenMai/ThietLap.asp?iStatus=Add" target="_parent"> Thêm mới</a> </strong></td>
      </tr>
<%
	set rs	=	Server.CreateObject("ADODB.recordset")
	sql = 	"Select * from SachTieuBieu"
	rs.open sql,con,1
	stt = 0
	Do while not rs.EOF
		stt=stt+1
%>	  
      <tr>
        <td style="<%=setStyleBorder(1,1,0,1)%>" align="center"><%=stt%></td>
        <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("TieuDe")%></td>
        <td style="<%=setStyleBorder(0,1,0,1)%>">
		<%
			if Clng(rs("KichHoat")) = 0 then
				Response.Write("<img src=""../images/icon-deactivate.gif"" width=""16"" height=""16"" border=""0"" align=""absmiddle"" alt=""Chưa hiển thị"">")
			else
				Response.Write("<img src=""../images/icon-activate.gif"" width=""16"" height=""16"" border=""0"" align=""absmiddle"" alt=""Đã hiển thị"">")
			end if
			
		%>		</td>
        <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
			<a href="EditSachTB.asp?iStatus=edit&ID=<%=rs("ID")%>" target="_parent"><img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" /></a>	
			<img src="../images/icon_closed_topic.gif" width="15" height="15" onClick="javascript: yn = confirm('Bạn có chắn chắn xóa ?'); if(yn) {winpopup('EditSachTB.asp','1&IDDel=<%=rs("ID")%>','100','100')}"/>			</td>
      </tr>
<%
		rs.movenext
	loop
%>	  
    </table></td>
  </tr>
  <tr>
    <td >&nbsp;</td>
  </tr>
</table>


<%Call Footer()%>

</body>
</html>