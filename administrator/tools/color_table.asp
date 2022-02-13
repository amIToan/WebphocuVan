<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_human")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>	
    <link href="../css/styles.css" rel="stylesheet" type="text/css">
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	=	"../../images/icons/icon_customer.jpg"
	Title_This_Page="Quản lý ->cập nhật bảng màu sản phẩm"
	Call header()		
	Call Menu()
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM ColorTable"
	rs.open sql,con,1
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	  <form action="color_table_update.asp"  target="_blank" name="colorList" method="post" enctype="multipart/form-data"> 
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0"  class=CTxtContent>
  <tr align="center">
    <td colspan="2"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">BẢNG MÀU SẢN PHẨM </font></strong><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
    </tr>

	<tr align="center" >
    <td width="388"  align="center">
	  <%
  i=1
  Do while not rs.eof
  %>
	<img src="<%=rs("PathColor")%>" border="0"  align="middle" alt="<%=rs("ID")%>"><span onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa màu này không?'); if(yn){winpopup('color_table_delete.asp','<%=rs("ID")%>&PathColor=<%=rs("PathColor")%>',300,150);}" style="cursor:pointer;">Xoá</span>&nbsp;&nbsp;&nbsp;
 <%
  if i = 3 then
  	Response.Write("<br>")
	i=0
  end if
  i=i+1
  rs.movenext
  Loop
  rs.close
  set rs=nothing
  %>	</td>
    <td width="80" >	</td>
  </tr>		

  
    <tr align="center" >
      <td height="26" colspan="2" align="left" class="CTieuDeNho" >&nbsp;</td>
    </tr>
    <tr align="center" >
      <td height="26" colspan="2" align="left" class="CTieuDeNho" background="../../images/BGMenu.jpg">Thêm mới</td>
    </tr>
    <tr align="center" >
    <td height="26" colspan="2" align="left">
	Mã màu:
	  <input name="IDAdd" type="text" class="CTextBoxUnder" value="" size="15" ><input name="PicColorFile" type="file" id="PicColorFile" size="17" ></td>
    </tr>
    <tr align="center" >
      <td height="26" colspan="2" align="left">&nbsp;</td>
    </tr>	
</table>
<center>
	<input name="submit" type="submit" id="submit" value="  Cập nhật  ">
    <input name="huy" type="button" id="huy" value=" Hủy  thao tác " onClick="javascript: window.location.reload();">
</center>
</form>
	  </td>
    </tr>
    <tr>
      <td background="../../images/T3.jpg" height="8"></td>
    </tr>
  </table>
<%Call Footer()%>
</body>
</html>