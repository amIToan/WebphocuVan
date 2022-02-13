<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
	<head>
		<title><%=PAGE_TITLE%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body>
<%
	img = "../../images/icons/Money-Pig-icon.jpg"	
	Call header()
	Call Menu()
%>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT top 1 * FROM exchange"
	rs.open sql,con,1
	if not rs.eof then
		vnd = rs("VND")
	end if
	rs.close
%>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="../../images/BGADM1.jpg" height="31"></td>
  </tr>
  <tr>
    <td background="../../images/BGADM2.jpg">
	<form action="update_exchange.asp"  target="_blank" name="YAHOOLIST" method="post">
<table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td class="author"><div align="center">ĐỊNH TỶ GIÁ USD - VNĐ </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="CTxtCommon" align="center">1 USD = 

<input name="vnd" type="text" class="CTextBoxUnder" id="vnd" size="15" maxlength="10" value="<%=Dis_str_money(vnd)%>" onKeyUp="javascript: DisMoneyThis(this);">      
VNĐ</td>
  </tr>
</table>
<br>
<center>
	<input name="submit1" type="submit" id="submit1" value="  Cập nhật  " >
    <input name="huy" type="button" id="huy" value=" Hủy  thao tác " onClick="javascript: window.location.reload();">
</center>
</form>
	</td>
  </tr>
  <tr>
    <td height="22" background="../../images/BGADM.jpg"></td>
  </tr>
</table>

<%Call Footer()%> 
</body>
</html>
