<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
email = Request.QueryString("Email")
sql = "SELECT * FROM SanphamUser WHERE emailCus = '"&email&"'"
	sql = sql & " ORDER BY SanphamUser_ID DESC"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
			If rs.eof Then
				%>
				<script language="javascript">
					top.resizeTo(500,150);
					top.moveTo(screen.width/2-250,screen.height/2-75);
					document.body.scroll = "yes";
				</script>
				<div>Không có đơn giao dịch nào thực hiện trong khoảng thời gian này.<div> <br>
				<div align="center">
					<span onClick="history.go(-1)" style="cursor:pointer; text-decoration:underline">Quay lại</span>
					&nbsp;&nbsp;&nbsp;
					<span onClick="window.close();" style="cursor:pointer; text-decoration:underline">Đóng</span>
				</div>
				
			<%Else%>
			
				<table border="0" width="100%">
				<tr><td>CÔNG TY CỔ PHẦN ĐỒNG XANH</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td align="center"><B>BẢNG KÊ THANH TOÁN TIỀN MUA HÀNG THỰC PHẨM METRO</B></td></tr>
				<tr><td align="center">Tên khách hàng: <%=getCustomerName(Email)%></td>
				</tr>
				<tr>
				  <td align="center">Email: <a href="/administrator/customer/send_mail.asp?email=<%= email %>"><%= email %></a></td>
				</tr>
				</table>
				<br>
				<table border="1" width="100%">
				<tr>
					<td align="center">STT</td>
					<td align="center">Mã GD</td>
					<td>Ngày đặt hàng</td>
					<td>Tổng giá trị</td>
					<td>Phát sinh nợ</td>
					<td>Số dư</td>
				</tr>
				<%
				j = 0
				CanDoi = 0 'cân đối dư và nợ
				TongCong = 0 ' tổng giá trị đơn hàng
				Do while not rs.eof
					Maso = rs("Maso")
					Ngay = rs("SanphamUser_Date")
					Total = rs("Total")
					If isNULL(Total) Then
						Total = "-Giao dịch chưa thành công-"
					Else
						TongCong = TongCong + Total
					End If
					Balance = rs("Balance")
					CanDoi = CanDoi + Balance
					j = j + 1
				%>
				 	<tr>
					<td align="center"><%=j%></td>
					<td align="center"><div onClick="window.location='/administrator/donhang/donhang_edit.asp?SanphamUser_ID=<%=rs("SanPhamUser_ID")%>'" style="cursor:pointer; color:#000099; text-decoration:underline"><%= rs("Maso") %></div></td>
					<td><%=Ngay%></td>
					<td><%=Total%></td>
					<td>
					<%If  Balance < 0 Then
						Response.Write(abs(Balance))
					Else
						Response.Write("0")
					End If
					%>
					</td>
					<td>
					<%If  Balance >= 0 Then
						Response.Write(abs(Balance))
					Else
						Response.Write("0")
					End If
					%>
					</td>
					<tr>
					<%
						rs.movenext
					Loop
					End If%>
			
				</table>
				<% If (CanDoi >0 ) Then 
						label = "Dư: " 
					Else
						label = "Nợ: "
					End If%>
				<div><strong>Tổng tiền hàng: <%= Dis_Str_Money(TongCONg) %> đ</strong>  <i>(<%=tienchu(TongCong)%>)</i></div>
				<div><strong><%= label %> <%= Dis_Str_Money(abs(CanDoi)) %> đ </strong><i>(<%=tienchu(abs(CanDoi))%>)</i></div>
</body>
</html>
 <%
 Function getCustomerName(email)
 	sqlN = "SELECT Name FROM Customer WHERE Email = '"&email&"'"
		Set rsN = Server.CreateObject("ADODB.Recordset")
	rsN.open sqlN,Con,1
	If not rsN.eof Then
		getCustomerName = rsN("Name")
	Else
		getCustomerName = "-NULL-"
	End If
	
 End function
 %>