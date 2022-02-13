<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_Datetime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
	CMND = Trim(Request.QueryString("param"))
	name_login=Trim(Request.QueryString("Name"))
	if CMND = "" then
		response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
				"window.close();" & vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline
	end if
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Lịch sử giao dịch tại XBOOK</title>
<link href="css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxbook128.png" width="157" height="57"></div></td>
    <td width="53%" rowspan="3" valign="bottom"><div align="center"><em>www.xbook.com.vn</em><br>
            <em>ĐT: <%=soDT%>  - Email: info@xbook.com.vn</em></div>      <div align="center"></div>      <div align="center"></div></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
  </tr>
  <tr>
    <td><div align="center"><em>ĐC: <%=dcVanPhong%> </em></div></td>
  </tr>
</table>
<br>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td align="center">
  <br>
        <font size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong>LỊCH SỬ giao DỊCH </strong></font>
		<br>
    </td>
  </tr>
  <tr>
    <td><u>Họ tên</u>: <%=name_login%></td>
  </tr>
  <tr>
    <td><u>Chứng minh thư</u>: <%=CMND%></td>
  </tr>
</table>
<br>
<%	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM V_SanPham_Donhang where CMND='"&CMND&"'"
	sql=sql+"ORDER BY SanPhamUser_ID DESC"
	rs.open sql,con,3
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</font></strong></td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	else

%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="CTxtContent">
    <td align="center" style="<%=setStyleBorder(1,1,1,1)%>"><strong>Số</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Địa chỉ giao hàng </strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Kiểm soát</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>giao hàng</strong> </td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tổng</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày đặt</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Trạng Thái </strong></td>
  </tr>
  <%
iMau=0
SoDH = 0
nTongTien =0
Do while not rs.eof 
SanPhamUser_ID=rs("SanPhamUser_ID")
SanPhamUser_Name=rs("SanPhamUser_Name")
SanPhamUser_Email=rs("SanPhamUser_Email")
SanPhamUser_Tell=rs("SanPhamUser_Tell")
SanPhamUser_Address=rs("SanPhamUser_Address")
SanPhamUser_Thoigian=rs("SanPhamUser_Thoigian")
SanPhamUser_Status=rs("SanPhamUser_Status")
Select case SanPhamUser_Status
	case 0
		StatusDonhang="Đơn hàng mới"
	case 1
		StatusDonhang="Đang xử lý"
	case 4
		StatusDonhang="Đợi sách"		
	case 2
		StatusDonhang="Đơn hàng đã xử lý"
	case 3
		StatusDonhang="Đơn hàng hủy bỏ"
End Select
SanPhamUser_Date=rs("SanPhamUser_Date")
giaoHang_Address=	rs("giaoHang_Address")
KSoat			=	getNhanVienFromID(rs("KiemSoat"))
giaoHang		=	getNhanVienFromID(rs("NhanVienID"))
%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%> class="CTxtContent">
    <td width="5%"align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
			
		%>
      <a href="../thongke/ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu"><%=strTemp%></a> </td>
    <td width="36%"align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=SanPhamUser_Address%></td>
    <td width="17%" align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%>. </td>
    <td width="16%" align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=giaoHang%>. </td>
    <td width="6%" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%
		iTien = 0
		iTien = LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,CMND))
		if iTien  >	DonhangMax then
			DonhangMax = iTien
		end if
		if 	iTien < DonHangMin then
			DonHangMin = iTien
		end if
		Response.Write(Dis_str_money(iTien)&"đ")
		nTongTien =	nTongTien + iTien
	%>
    </td>
    <td width="5%" align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Hour(SanPhamUser_Date)%>h<%=Minute(SanPhamUser_Date)%>'<br>
        <%=Day(SanPhamUser_Date)%>/<%=Month(SanPhamUser_Date)%>/<%=Year(SanPhamUser_Date)%></td>
    <td width="10%" style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=StatusDonhang%></td>
  </tr>
  <%
SoDH = SoDH+1
rs.movenext
Loop%>
</table>
<p>
  <%	end if 'if not rs.eof then
	rs.close
	set rs=nothing
%>	
</p>

<table width="400" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
    <td width="38%">Tổng đơn hàng:</td>
    <td width="62%"><%=SoDH%></td>
  </tr>
  <tr>
    <td>Tổng tiền:</td>
    <td><%=Dis_str_money(nTongTien)%>đ</td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
<p align="center" >
  <input type="button" name="Button" value="Đóng lại" onClick="javascript:window.close();">
</p>
</body>
</html>
