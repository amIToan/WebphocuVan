<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%SanPhamUser_ID = GetNumeric(Request.querystring("SanPhamUser_ID"),0)%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <link href="../../include/style.css" rel="stylesheet" type="text/css">
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	 <table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
    <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
    <em>ĐT: <%=soDT%>  - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%> </em></div></td>
  </tr>
</table>
<br>
<br>
	 <center>
	   <font class="CFontVerdana10"> ĐƠN HÀNG CHI TIẾT</font><br><%=Day(now)%>/<%=Month(Now)%>/<%=year(Now)%>
	 </center>
<%
		sql="SELECT	TOP 1 * FROM  SanPhamUser WHERE SanPhamUser_ID="&SanPhamUser_ID 
		Set rs=server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		if not rs.eof then
			SanPhamUser_Name=rs("SanPhamUser_Name")
			SanPhamUser_Email=rs("SanPhamUser_Email")
			SanPhamUser_Tell=rs("SanPhamUser_Tell")
			SanPhamUser_Address=rs("SanPhamUser_Address")
			SanPhamUser_Thoigian=rs("SanPhamUser_Thoigian")
			SanPhamUser_Status=rs("SanPhamUser_Status")	
			strCMND			=	rs("CMND")	
			GiaoHang_name 	= 	rs("GiaoHang_name")
			GiaoHang_Email	= 	rs("GiaoHang_Email")
			GiaoHang_Tel	=	rs("GiaoHang_Tel")		
			GiaoHang_Address=	rs("GiaoHang_Address")
			GiaoHang_times	= 	rs("GiaoHang_times")
			GiaoHang_YeuCauK= 	rs("GiaoHang_YeuCauK")
			NhanVienID		=	rs("NhanVienID")
			KiemSoat		=	rs("KiemSoat")
		 end if
		 rs.close
		 set rs=nothing
%>	
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
          <tr>
            <td colspan="2" align="center" ><u>ĐỊA CHỈ GIAO HÀNG </u></td>
            <td colspan="2" align="center"><u>ĐỊA CHỈ THANH TOÁN</u></td>
          </tr>
          <tr>
            <td colspan="2" align="right" >&nbsp;</td>
            <td colspan="2" align="right" >&nbsp;</td>
          </tr>
          <tr>
            <td width="11%" align="right" >Họ và Tên :</td>
            <td width="37%" ><b><%=SanPhamUser_Name%></b> </td>
            <td width="13%" align="right" >Họ và Tên:</td>
            <td width="39%" ><strong><%=GiaoHang_name%></strong></td>
          </tr>
          <tr>
            <td align="right" >Email :</td>
            <td ><strong><%=SanPhamUser_Email%></strong></td>
            <td align="right" >Email:</td>
            <td ><strong><%=GiaoHang_Email%></strong></td>
          </tr>
          <tr>
            <td align="right" >Điện thoại:</td>
            <td ><strong><%=SanPhamUser_Tell%></strong></td>
            <td align="right" >Điện thoại: </td>
            <td ><strong><%=GiaoHang_Tel%></strong></td>
          </tr>
          <tr>
            <td align="right" >Địa chỉ :</td>
            <td ><strong><%=SanPhamUser_Address%></strong></td>
            <td align="right" >Địa chỉ:</td>
            <td ><strong><%=GiaoHang_Address%></strong></td>
          </tr>
          
          <tr>
            <td colspan="4" align="left" >Yêu cầu khác:<%=GiaoHang_YeuCauK%></td>
          </tr>
</table>
		<br>
<table width="100%" border="0"  align="center" cellpadding="0"  cellspacing="0" bordercolor="#CCCCCC">
		  <tr>
			<td width="7%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
			<td width="10%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Mã SP </td>
			<td width="41%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Tên SP </td>
			<td width="9%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Giá bìa </td>
			<td width="5%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">SL </td>
			<td width="9%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Thu </td>
			<td width="9%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Chi</td>
			<td width="10%" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Lợi nhuận </td>
	      </tr>
	<%
	stt= 1
	TTien = 0
	sql = "SELECT NewsID,SanPham_User_ID,SanPham_ID,idsanpham,Title,Tacgia,SanPham_Giabia,SanPham_Soluong,SanPham_Gia  "
	sql= sql + " FROM SanPham_User INNER JOIN SanPhamNhap ON SanPham_User.SanPham_ID = SanPhamNhap.NewsID "
	sql= sql + " WHERE SanPhamUser_ID = '"& SanPhamUser_ID &"' and re_newsid = 0"
	set rsProduct=Server.CreateObject("ADODB.Recordset")
	rsProduct.open sql,Con,1 
	TongThu = 0
	TongChi = 0
	Do while not rsProduct.eof
	%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
			<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProduct("idsanpham")%></td>
			<td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=LCase(rsProduct("Title"))%></td>
			<td align="right" style="<%=setStyleBorder(0,1,0,1)%>">
			<%=Dis_str_money(rsProduct("SanPham_Giabia"))%></td>
	        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProduct("SanPham_Soluong")%> </td>
          <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">
<%			
			TongThu  = TongThu +  rsProduct("SanPham_Gia")*rsProduct("SanPham_Soluong")
		  Response.Write(Dis_str_money(rsProduct("SanPham_Gia")*rsProduct("SanPham_Soluong")))
%>		  </td>
	<td align="right" style="<%=setStyleBorder(0,1,0,1)%>">
		<%
		TongChi	= TongChi + GiaNhapKho(rsProduct("SanPham_User_ID"))*GetTotalXuatKho(rsProduct("SanPham_User_ID"))
		Response.Write(Dis_str_money(GiaNhapKho(rsProduct("SanPham_User_ID"))*GetTotalXuatKho(rsProduct("SanPham_User_ID"))))
		%>		</td>
	<td align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
	if GetTotalXuatKho(rsProduct("SanPham_User_ID"))= rsProduct("SanPham_Soluong")then
		iLN =(Clng(rsProduct("SanPham_Gia")) - GiaNhapKho(rsProduct("SanPham_User_ID")))*GetTotalXuatKho(rsProduct("SanPham_User_ID"))
		TTien =TTien+iLN
		Response.Write(Dis_str_money(iLN))
	else
		Response.Write("Chưa đủ SLượng")
	end if
	%>	</td>
		</tr>
	<%
		stt = stt +1
		rsProduct.movenext
	loop
	%>
	<tr class="CTxtContent">
	  <td colspan="5" align="right" class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>">Tổng:</td>
	  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(TongThu)%></td>
	  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(TongChi)%></td>
	  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(TongThu-TongChi)%></td>
  </tr>
	<tr class="CTxtContent">
		  <td colspan="3" align="right" class="CTxtContent">Cước bưu điện thu khách hàng: </td>
			<td align="right"><%=Dis_str_money(GetCuocBuuDienID(SanPhamUser_ID))%></td>
			<td colspan="2" align="right">Chi:</td>
			<td align="right"><%=Dis_str_money(GetCuocBuuDienThucID(SanPhamUser_ID))%></td>
			<td align="right"><%=Dis_str_money(GetCuocBuuDienID(SanPhamUser_ID)-GetCuocBuuDienThucID(SanPhamUser_ID))%></td>
  </tr>
	<tr class="CTxtContent">
	  <td colspan="3" align="right" class="CTxtContent">Khoản thu khác: </td>
	  <td align="right"><span style="<%=setStyleBorder(0,0,0,1)%>"><%=Dis_str_money(GetThuKhac(SanPhamUser_ID))%></span></td>
	  <td colspan="2" align="right">Chi: </td>
	  <td align="right"><%=Dis_str_money(GetChiKhac(SanPhamUser_ID))%></td>
	  <td align="right"><%=Dis_str_money(GetThuKhac(SanPhamUser_ID) - GetChiKhac(SanPhamUser_ID))%></td>
  </tr>
	<tr class="CTxtContent">
	  <td colspan="7" align="right">Phí vận chuyển: </td>
	  <td align="right"><%=Dis_str_money(GetPhiVanChuyen(SanPhamUser_ID))%></td>
  </tr>
	<tr class="CTxtContent">
	  <td colspan="7" align="right">&nbsp;</td>
	  <td align="right" style="<%=setStyleBorder(0,0,0,1)%>">&nbsp;</td>
  </tr>
	<tr>
	  <td colspan="7" align="right" class="CFontVerdana10">Tổng lợi nhuận:</td>
	  <td align="right">
	  <%
	  fDuBuuDien 	= 	GetCuocBuuDienID(SanPhamUser_ID)-GetCuocBuuDienThucID(SanPhamUser_ID)
	  fChiPhiKhac	=	GetThuKhac(SanPhamUser_ID) - GetChiKhac(SanPhamUser_ID)
	  TTien = TTien + GetPhiVanChuyen(SanPhamUser_ID) + fDuBuuDien + fChiPhiKhac
	  Response.Write(Dis_str_money(TTien))
	  %>	  </td>
  </tr>
</table>
<br>
<%
	  if GetThuKhac(SanPhamUser_ID) > 0 then
	  	Response.Write("- <b>Lưu ý</b>: *Khoản thu khác do "&GetGhiChu(SanPhamUser_ID)&"<br>")
	  end if
	  %>
	  <%if TRIM(TBPhiVanChuyen) <> "1" and trim(TBPhiVanChuyen) <> "0" then Response.Write("- "&TBPhiVanChuyen) end if%><br>
	  <%
	  if TRIM(SanPhamUser_Thoigian) <> "Không quá 24 tiếng" then 
	  	Response.Write("- "&SanPhamUser_Thoigian&"<br>")
	   else
		Response.Write("- <b>Thời gian giao hàng:</b> "&SanPhamUser_Thoigian&"<br>")
	  end if
	  if TRIM(GiaoHang_times) <> "Không quá 24 tiếng" then 
	  	Response.Write("- "&GiaoHang_times)
	  end if
		%><br>
<br>
<table width="100%" border="0" cellspacing="2" cellpadding="2" class="CTxtContent">
  <tr>
    <td width="48%" align="center"><strong>Kiểm soát viên </strong></td>
    <td width="52%" align="center"><strong>Giao hàng </strong></td>
  </tr>
  <tr>
    <td align="center" class="CFontVerdana10"><%=GetNameNV(KiemSoat)%></td>
    <td align="center" class="CFontVerdana10"><%=GetNameNV(NhanVienID)%></td>
  </tr>
</table>
</body>
</html>
