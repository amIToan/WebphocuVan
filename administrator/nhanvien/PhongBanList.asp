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
	Title_This_Page="Quản lý ->cập nhật chức danh"
	Call header()		
	Call Menu()
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM PhongBan"
	rs.open sql,con,1
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	  <form action="PhongbanUpdate.asp?action=update"  target="_parent" name="PhongbanList" method="post"> 
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0"  class=CTxtContent>
  <tr align="center">
    <td width="55"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">STT </font></strong></td> 
    <td width="388"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Phòng ban </font></strong></td>
    <td width="80" ><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
  </tr>
  <%
  i=1
  Do while not rs.eof
  %>
	<tr align="center" >
    <td width="55"><%=rs("PhongID")%><input name="IDPhongBan<%=i%>" type="hidden" class="CTextBoxUnder" value="<%=rs("PhongID")%>" size="3" onBlur="checkIsNumber(this)" ></td> 
    <td width="388" align="left">
	<input name="txtPhongban<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Description")%>" size="50">
	</td>
    <td width="80" >
	<span onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa Phòng này không?'); if(yn){ winpopup('PhongbanUpdate.asp','0&IDPhongBan=<%=rs("PhongID")%>&action=del',300,150);}" style="cursor:pointer;">Xoá</span>
</td>
  </tr>		
  <%
  i=i+1
  rs.movenext
  Loop
  rs.close
  set rs=nothing
  %>
  
  <tr align="center" >
    <td width="55" height="26"><%=i%><input name="IDPhongBan<%=i%>" type="hidden" class="CTextBoxUnder" value="<%=i%>" size="3" onBlur="checkIsNumber(this)"></td> 
    <td width="388" align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtPhongban<%=i%>" type="text" class="CTextBoxUnder" value="" size="35" >
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
<script language="javascript">
function CheckNXBNew()
{
	var i
	i= document.NXBLIST.iCount.value;
	str = 'document.NXBLIST.txtNXB&i.value'
	eval(str) = ''
}
</script>