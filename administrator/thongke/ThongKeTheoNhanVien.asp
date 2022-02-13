<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/include/funcNotData.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language="javascript">
top.moveTo(screen.width/2-400,screen.length/2-400)
</script>
<title>Thống kê theo nhân viên</title>
<style type="text/css">
.GiaodichTitle
{
background:#CCCCCC; color:#000066; font-weight:bold;
}
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
NhanvienID = Request.QueryString("NhanvienID") 
sql = "SELECT * FROM Nhanvien nv JOIN Branch b ON nv.BranchID = b.BranchID WHERE NhanvienID = " & NhanvienID
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,1
If rs.eof  Then
	TenNhanVien = ""
	TenChiNhanh = ""
Else
	TenNhanVien = rs("Ho_Ten")
	TenChiNhanh = rs("BranchName")
End If
FromDate = CDate(Request.QueryString("D1"))
ToDate = CDate(Request.QueryString("D2"))
Set rs = nothing

NhanVienID = Clng(NhanVienID)
sql = "SELECT * FROM SanphamUser WHERE (GiaoHang = " & NhanVienID & " OR KiemSoat = " & NhanVienID & ")"
sql = sql & " AND (SanphamUser_Date BETWEEN '"&DateAdd("d",-1,FromDate)&"' AND '"&DateAdd("d",1,ToDate)&"')"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,1
If not rs.eof Then
%>
<table width="100%" border="0">
  <tr>
    <td colspan="3">CÔNG TY CỔ PHẦN ĐỒNG XANH </td>
  </tr>
  <tr>
    <td width="7%">&nbsp;</td>
    <td width="35%">&nbsp;</td>
    <td width="58%">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" align="center"><strong>BÁO CÁO CA BÁN HÀNG THỰC PHẨM THEO NHÂN VIÊN</strong> </td>
  </tr>
  <tr>
    <td colspan="3" align="center">Họ và tên nhân viên: <%= TenNhanVien %> - Chi nhánh: <%= TenChiNhanh %> </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2" align="center">Từ ngày: <%=FromDate  %> &nbsp;-&nbsp; Đến ngày: <%= ToDate %></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<p>KIỂM SOÁT VIÊN </p>
<table width="100%" border="1">
  <tr align="center">
    <td width="6%">STT</td>
    <td width="10%">Mã số GD </td>
    <td width="23%">Điểm chuyển phát hàng </td>
    <td width="12%">Tiền hàng </td>
    <td width="13%">Dư/Nợ</td>
  </tr>
  <tr>
    <td>I</td>
    <td colspan="4" class="GiaodichTitle">Giao dịch đã hoàn thành </td>
  </tr>
  <%
  i = 0
  total = 0
  balance = 0
  Do While not rs.eof
  		If rs("KiemSoat") = NhanVienID AND rs("SanphamUser_Status") = 2 Then 'nếu là kiểm soát viên và giao dịch đã hoàn thành
			i = i + 1
  %>
  <tr>
    <td><%= i %></td>
    <td><div onClick="window.location='/administrator/donhang/donhang_edit.asp?SanphamUser_ID=<%=rs("SanPhamUser_ID")%>'" style="cursor:pointer"><%= rs("Maso") %></div></td>
    <td><%= rs("SanphamUser_Address") %></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Total")) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Balance")) %> đ</div></td>
  </tr>
  <% 
  			total = total + rs("Total")
			balance = balance + rs("Balance")
  		End If
  	rs.movenext
	total1 = total
	balance1 = balance
  Loop %>
  <tr>
    <td>&nbsp;</td>
    <td>Cộng:</td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance) %> đ</div></td>
  </tr>
  <tr>
    <td>II</td>
    <td colspan="4" class="GiaodichTitle">Giao dịch chưa hoàn thành </td>
  </tr>
  <%
  i = 0
  total = 0
  balance = 0
  rs.movefirst
  Do While not rs.eof
  		If rs("KiemSoat") = Clng(NhanvienID) AND rs("SanphamUser_Status")= 1 Then 'nếu là KIỂM SOÁT VIÊN và giao dịch CHƯA hoàn thành
			i = i + 1
  %>
  <tr>
    <td><%= i %></td>
    <td><div onClick="window.location='/administrator/donhang/donhang_edit.asp?SanphamUser_ID=<%=rs("SanPhamUser_ID")%>'" style="cursor:pointer"><%= rs("Maso") %></div></td>
    <td><%= rs("SanphamUser_Address") %></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Total")) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Balance")) %> đ</div></td>
  </tr>
  <% 
  			total = total + rs("Total")
			balance = balance + rs("Balance")
  		End If
  	rs.movenext
	total2 = total
	balance2 = balance
  Loop %>
  <tr>
    <td>&nbsp;</td>
    <td>Cộng:</td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance) %></div></td>
  </tr>
  <tr>
    <td colspan="5">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
	<td><strong>Tổng cộng: </strong></td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total1 + total2) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance1 + balance2) %> đ</div></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>NHÂN VIÊN GIAO HÀNG </p>
<table width="100%" border="1">
  <tr align="center">
    <td width="6%">STT</td>
    <td width="10%">Mã số GD </td>
    <td width="23%">Điểm chuyển phát hàng </td>
    <td width="12%">Tiền hàng </td>
    <td width="13%">Dư/Nợ</td>
  </tr>
  <tr>
    <td>I</td>
    <td colspan="4" class="GiaodichTitle">Giao dịch đã hoàn thành </td>
  </tr>
  <%
  i = 0
  total = 0
  balance = 0
  rs.movefirst
  Do While not rs.eof
  		If rs("GiaoHang") = NhanVienID AND rs("SanphamUser_Status") = 2 Then 'nếu là kiểm soát viên và giao dịch đã hoàn thành
			i = i + 1
  %>
  <tr>
    <td><%= i %></td>
    <td><div onClick="window.location='/administrator/donhang/donhang_edit.asp?SanphamUser_ID=<%=rs("SanPhamUser_ID")%>'" style="cursor:pointer"><%= rs("Maso") %></div></td>
    <td><%= rs("SanphamUser_Address") %></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Total")) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Balance")) %> đ</div></td>
  </tr>
  <% 
  			total = total + rs("Total")
			balance = balance + rs("Balance")
  		End If
  	rs.movenext
	total1 = total
	balance1 = balance
  Loop %>
  <tr>
    <td>&nbsp;</td>
    <td>Cộng:</td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance) %> đ</div></td>
  </tr>
  <tr>
    <td>II</td>
    <td colspan="4" class="GiaodichTitle">Giao dịch chưa hoàn thành </td>
  </tr>
  <%
  i = 0
  total = 0
  balance = 0
  rs.movefirst
  Do While not rs.eof
  		If rs("GiaoHang") = Clng(NhanvienID) AND rs("SanphamUser_Status")= 1 Then 'nếu là KIỂM SOÁT VIÊN và giao dịch CHƯA hoàn thành
			i = i + 1
  %>
  <tr>
    <td><%= i %></td>
    <td><div onClick="window.location='/administrator/donhang/donhang_edit.asp?SanphamUser_ID=<%=rs("SanPhamUser_ID")%>'" style="cursor:pointer"><%= rs("Maso") %></div></td>
    <td><%= rs("SanphamUser_Address") %></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Total")) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(rs("Balance")) %> đ</div></td>
  </tr>
  <% 
  			total = total + rs("Total")
			balance = balance + rs("Balance")
  		End If
  	rs.movenext
	total2 = total
	balance2 = balance
  Loop %>
  <tr>
    <td>&nbsp;</td>
    <td>Cộng:</td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance) %></div></td>
  </tr>
  <tr>
    <td colspan="5">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><strong>Tổng cộng: </strong></td>
    <td>&nbsp;</td>
    <td><div align="right"><%= Dis_Str_Money(total1 + total2) %> đ</div></td>
    <td><div align="right"><%= Dis_Str_Money(balance1 + balance2) %> đ</div></td>
  </tr>
</table>
<p>&nbsp;</p>
<br>
<div id="printButton" align="center"><input type="button" value="    In    " onClick="document.getElementById('printButton').style.display = 'none'; window.print();"> </div>
<% Else %>
Không có dữ liệu!
<% End If %>
<p>&nbsp;</p>
</body>
</html>