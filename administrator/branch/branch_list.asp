<%session.CodePage=65001%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_order_input")
if f_permission < 3 then
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
	Title_This_Page="Quản lý -> Nhập liệu nhà cung cấp."
	Call header()
	Call Menu()
%>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM Branch ORDER BY BranchName"
	rs.open sql,con,1
%>
<br>
<form action="branch_update.asp?action=update"  name="BranchLIST" method="post">
<table width="80%" border="1" align="center" cellpadding="0" cellspacing="0"  class=normal>
  <tr align="center">
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">STT </font></strong></td> 
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
      chi nhánh </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Địa chỉ </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Điện thoại </font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Di động </font></strong></td>
	<td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
  </tr>
  <%
  i=1
  Do while not rs.eof
  %>
	<tr align="center" >
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=i%> </font>
	<input name="BranchID<%=i%>" type="hidden" value="<%=rs("BranchID")%>"></td> 
	
    <td align="left"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtBranchName<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("BranchName")%>" size="20">
	</font></strong></td>

    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtAddress<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Address")%>" size="40">
    </font></td>

    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtTel<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Tel")%>" size="20"> 
	</font></td>

    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtMobile<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("mobile")%>" size="20">
    </font></td>

	<td>
	<%if Session("iQuanTri") = 1 then %>
	<span onClick="javascript:window.open('branch_update.asp?branchID=<%=rs("BranchID")%>&action=del', 'mywindow',10,10)">Xoá</span>
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
	<input name="txtBranchName<%=i%>" type="text" class="CTextBoxUnder" value="" size="20" >
	</font></strong></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtAddress<%=i%>" type="text" class="CTextBoxUnder" value="" size="40">
    </font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtTel<%=i%>" type="text" class="CTextBoxUnder" value="" size="20"> 
	</font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
	<input name="txtMobile<%=i%>" type="text" class="CTextBoxUnder" value="" size="20">
    </font></td>
	<td>&nbsp;</td>
  </tr>	
</table>
<center>
	<input type="hidden" name="iCount" value="<%=i%>">
	<input type="hidden" name="action" value="update">
	<input name="submit1" type="button" id="submit1" value="  Cập nhật  " onClick="checkInput();">
    <input name="huy" type="button" id="huy" value=" Hủy  thao tác " onClick="javascript: window.location.reload();">
</center>
</form>
<%Call Footer()%>
</body>
</html>
<script language="javascript">
function checkInput()
{
	i = document.BranchLIST.iCount.value;
	for(j = 1;j < i; j++)
	{
		str = "tel = document.BranchLIST.txtTel"+j+".value";
		eval(str);
		if (isNaN(tel))
		{
			str2 = "BranchName = document.BranchLIST.txtBranchName"+j+".value";
			eval(str2);
			alert("Kiểm tra lại số điện thoại của chi nhánh " + BranchName + ".");
			str3 = "document.BranchLIST.txtTel"+j+".select();"
			eval(str3);
			return;
		}
		str = "mobile = document.BranchLIST.txtMobile"+j+".value";
		eval(str);
		if (isNaN(mobile))
		{
			str2 = "BranchName = document.BranchLIST.txtBranchName"+j+".value";
			eval(str2);
			alert("Kiểm tra lại số điện thoại của chi nhánh " + BranchName + ".");
			str3 = "document.BranchLIST.txtMobile"+j+".select();"
			eval(str3);
			return;
		}
	}
	document.BranchLIST.submit();
}
</script>