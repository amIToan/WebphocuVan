<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_accounting")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
	<head>
		<title><%=PAGE_TITLE%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	=	"../../images/icons/bank_48.jpg"
	Title_This_Page="Ngân hàng - > Cậo nhật danh sách"
	Call header()
	Call Menu()
%>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM Bank ORDER BY BankName"
	rs.open sql,con,1
%>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td ></td>
  </tr>
  <tr>
    <td >
	<form action="Bank_update.asp?action=update"   name="BankLIST" method="post">
<br>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1"  class=normal>
  <tr align="center">
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">STT </font></strong></td> 
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
      ngân hàng </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Địa chỉ </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Điện thoại </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
  </tr>
  <%
  i=1
  Do while not rs.eof
  %>
	<tr align="center" <%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=i%> </font>
	<input name="BankID<%=i%>" type="hidden" value="<%=rs("BankID")%>"></td> 
	
    <td align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtBankName<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("BankName")%>" size="25">
	</font></strong></td>

    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtAddress<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Address")%>" size="15">
    </font></td>

    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtTel<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Tel")%>" size="50"> 
	</font></td>

    <td>
	<%if Session("iQuanTri")= 1 then%>
	<span onClick="javascript:window.open('Bank_update.asp?BankID=<%=rs("BankID")%>&action=del', 'mywindow',100,100)">Xoá</span>
	<%end if%>
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
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=i%> </font></td> 
    <td align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtBankName<%=i%>" type="text" class="CTextBoxUnder" value="" size="25" >
	</font></strong></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtAddress<%=i%>" type="text" class="CTextBoxUnder" value="" size="15">
    </font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtTel<%=i%>" type="text" class="CTextBoxUnder" value="" size="50"> 
	</font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">	</td>
  </tr>	
</table>
<center>
	<input type="hidden" name="iCount" value="<%=i%>">
	<input type="hidden" name="action" value="update">
	<input name="submit1" type="button" id="submit1" value="  Cập nhật  " onClick="checkInput();">
    <input name="huy" type="button" id="huy" value=" Hủy  thao tác " onClick="javascript: window.location.reload();">
</center>
</form>
	</td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>

<%Call Footer()%> 
</body>
</html>
<script language="javascript">
function checkInput()
{
	i = document.BankLIST.iCount.value;
	for(j = 1;j < i; j++)
	{
		str = "tel = document.BankLIST.txtTel"+j+".value";
		eval(str);
		if (isNaN(tel))
		{
			str2 = "bankName = document.BankLIST.txtBankName"+j+".value";
			eval(str2);
			alert("Kiểm tra lại số điện thoại của ngân hàng " + bankName + ".");
			str3 = "document.BankLIST.txtTel"+j+".select();"
			eval(str3);
			return;
		}
	}
	document.BankLIST.submit();
}
</script>