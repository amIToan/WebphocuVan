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
	Title_This_Page="Quản lý ->cập nhật chức danh"
	Call header()		
	Call Menu()
%>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="right" valign="top"> 
    <td height="25"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
		<a href="javascript: winpopup('cat_chooselang.asp','<%=lang%>',220,120);"> 
      Chọn Ngôn ngữ</a></strong></font> 
    </td>
  </tr>
</table>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM ChucDanh"
	rs.open sql,con,1
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	  <form action="ChucDanhUpdate.asp?action=update"  target="_parent" name="ChucDanhList" method="post">
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0"  class=CTxtContent>
  <tr align="center">
    <td width="55"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">STT </font></strong></td> 
    <td width="388"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
      chức danh </font></strong></td>
    <td width="80" ><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
  </tr>
  <%
  i=1
  Do while not rs.eof
  %>
	<tr align="center" >
    <td width="55"><%=rs("ChucDanhID")%><input name="IDChucDanh<%=i%>" type="hidden" class="CTextBoxUnder" value="<%=rs("ChucDanhID")%>" size="3" onBlur="checkIsNumber(this)"></td> 
    <td width="388" align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtChucDanh<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Description")%>" size="50">
	</font></strong></td>
    <td width="80" ><span onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa Phòng này không?'); if(yn){ winpopup('PhongbanUpdate.asp','0&IDChucDanh=<%=rs("ChucDanhID")%>&action=del',300,150);}" style="cursor:pointer;">Xoá</span></td>
  </tr>		
  <%
  i=i+1
  rs.movenext
  Loop
  rs.close
  set rs=nothing
  %>
  
  <tr align="center" >
    <td width="55" height="26"><%=i%><input name="IDChucDanh<%=i%>" type="hidden" class="CTextBoxUnder" value="<%=i%>" size="3" onBlur="checkIsNumber(this)"></td> 
    <td width="388" align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtChucDanh<%=i%>" type="text" class="CTextBoxUnder" value="" size="35" >
	</font></strong></td>
    <td width="80" ></td>
  </tr>	
</table>
<center>
	<input type="hidden" name="iCount" value="<%=i%>">
	<input type="hidden" name="action" value="update">
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