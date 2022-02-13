<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%

Ngay1			=	GetNumeric(Request.form("Ngay1"),0)
Thang1			=	GetNumeric(Request.form("Thang1"),0)
Nam1			=	GetNumeric(Request.form("Nam1"),0)
Ngay2			=	GetNumeric(Request.form("Ngay2"),0)
Thang2			=	GetNumeric(Request.form("Thang2"),0)
Nam2			=	GetNumeric(Request.form("Nam2"),0)
iOrderBy		=  	Clng(Request.Form("RaOderBy"))
ctXuatKho		=	GetNumeric(Request.form("ctXuatKho"),0)
ctNhapKho		=	GetNumeric(Request.form("ctNhapKho"),0)
ctTonKho		=	GetNumeric(Request.form("ctTonKho"),0)
ctTamUng		=	GetNumeric(Request.form("ctTamUng"),0)
ctxseo			=	GetNumeric(Request.form("ctxseo"),0)
ctHuyBo			=	GetNumeric(Request.form("ctDHHuy"),0)



FromDate=Thang1&"/"&Ngay1&"/"&Nam1
ToDate=Thang2&"/"&Ngay2&"/"&Nam2
		
%>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
    <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
    <em>ĐT: (04) 2922.446 - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%>   </em></div></td>
  </tr>
</table>
<br>
  <div align="center"class="author">
    <div align="center"><strong>HẠCH TOÁN </strong></div>
  </div>
  <center> Từ ngày <%=Ngay1%>/<%=Thang1%>/<%=Nam1%> Đến <%=Ngay2%>/<%=Thang2%>/<%=Nam2%></center>
<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM SanPhamUser " 
	sql=sql + " where (SanPhamUser_Status= 2 or SanPhamUser_Status=14 or SanPhamUser_Status=1 or SanPhamUser_Status=6 or SanPhamUser_Status = 15) and "
	sql=sql + "  (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
	sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
	rs.open sql,con,3
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	end if
	sqlnv	= "Select Count(NhanVienID) as iCount From NhanVien"
	set rsnv = Server.CreateObject("ADODB.recordset")
	rsnv.open sqlnv,con,1
	iCountNV	=rsnv("iCount")
	Redim arNhanVienID(iCountNV)
	redim arNhanVienValues(iCountNV)
	redim arNhanVienChi(iCountNV)
	redim arNhanVienTamUng(iCountNV)
	redim arNhanVienTTamUng(iCountNV)
		
	set rsnv = nothing
	sqlnv	= "Select NhanVienID From NhanVien"
	set rsnv = Server.CreateObject("ADODB.recordset")
	rsnv.open sqlnv,con,1
	h = 0
	do while not rsnv.eof
		arNhanVienID(h) = rsnv("NhanVienID")
		h= h +1
		rsnv.movenext
	loop
	set rsnv = nothing	
%>
<div align="center" class="author">
  <div align="left"><strong>XUẤT KHO</strong> </div>
</div>
<%if ctXuatKho = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr> 
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td width="39%" align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tên</b></td>
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Giao hàng </td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Kiểm soát</b> </div></td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Thu tiền</b> </div></td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Chi </td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Thu</td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Dư</td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Ngày đặt</b> </td>
  </tr>
<%end if%>  
<%
iMau=0
STT = 0
fTongTienXuat =0
fTongTienThu =0

Do while not rs.eof 
SanPhamUser_ID		=	rs("SanPhamUser_ID")
SanPhamUser_Name	=	rs("SanPhamUser_Name")
SanPhamUser_Email	=	rs("SanPhamUser_Email")
SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
SanPhamUser_Address	=	rs("SanPhamUser_Address")
NgayXuLy	=	rs("NgayXuLy")
strCMND				=	rs("CMND")
KSoat				=	getNhanVienFromID(rs("KiemSoat"))
NVGiaoHang			=	getNhanVienFromID(rs("NhanVienID"))	
NVThutien			=	getNhanVienFromID(rs("NVThutienID"))	
if KSoat = "" then
	KSoat="&nbsp;"
end if
if NVGiaoHang = "" then
	NVGiaoHang="&nbsp;"
end if

	sql =       "SELECT XuatKho.SoLuong,Product.Price,Product.VAT "
	sql = sql + " FROM SanPham_User INNER JOIN SanPhamUser ON SanPham_User.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID"
	sql = sql + " INNER JOIN XuatKho ON SanPham_User.SanPham_User_ID = XuatKho.SanPham_User_ID"
	sql = sql + " INNER JOIN Product ON XuatKho.ProductID = Product.ProductID"
	sql = sql + " WHERE SanPhamUser.SanPhamUser_ID = '"&SanPhamUser_ID&"' and SanPham_User.re_newsid = 0"
	set rss = Server.CreateObject("ADODB.recordset")
	rss.open sql,con,3
	fTongXuat = 0
	Do while not rss.eof
		fDongia		=	rss("Price") + 	rss("Price")*rss("VAT")/100
		fDongia		=	fDongia*rss("SoLuong")
		fTongXuat 	=	fTongXuat + fDongia 
		rss.movenext
	loop
	set rss= nothing 
	fTongXuat =fTongXuat + GetCuocBuuDienThucID(SanPhamUser_ID) + GetChiKhac(SanPhamUser_ID)
	fTongTienXuat = fTongTienXuat + fTongXuat 
	
	iTien 	= 	0
	iTien 	= 	TongTienTrenDonHang(SanPhamUser_ID,strCMND)
	fTongTienThu =	fTongTienThu + iTien
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThutienID") then
			arNhanVienValues(h) = arNhanVienValues(h) + iTien
		end if
	next
	
%>
<%if ctXuatKho = 1 then%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td width="3%"align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu"><%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a></td>
    <td align="left" valign="middle" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%><br>
	<font class="CSubTitle">	
	<i>Điện thoại</i>: <%=SanPhamUser_Tell%><br>
	<i>Địa chỉ</i>: <%=SanPhamUser_Address%></font>	</td>
    <td width="12%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=NVGiaoHang%></td>
	<td width="13%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%></td>
	<td width="12%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=NVThutien%></td>
	<td width="5%"  style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(fTongXuat)%></td>
	<td width="5%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(iTien)%>	</td>
	<td width="5%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(iTien-fTongXuat)%>	</td>
	<td width="6%" style="<%=setStyleBorder(0,1,0,1)%>">
		<%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%>
	</td>
  </tr>

<%
iMau=iMau+1
end if
stt=stt + 1
rs.movenext
Loop
%>
<%if ctXuatKho = 1 then%>
</table>
<%end if%>
<%	rs.close
	set rs=nothing
%>	
<br>
    <table width="700" border="0" cellpadding="2" cellspacing="2">
      <tr>
        <td colspan="3" class="CTxtContent" >Tổng số lượng: <%=stt%> đơn hàng
        <div align="left"></div></td>
      </tr>		  
      <tr>
        <td width="161" align="left" class="CTxtContent">Tổng tiền thu:</td>
        <td width="141" align="left" valign="top" class="CTxtContent" style="<%=setStyleBorder(0,1,0,0)%>"><b><%=Dis_str_money(fTongTienThu)%></b>		</td>
        <td width="378" align="left" class="CTxtContent">
		<table border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
<%
	for h = 0 to iCountNV-1 
	 if arNhanVienValues(h) <> "" and  arNhanVienValues(h) <> 0 then
%>			<tr>
				<td align="right">
				<b><%=getNhanVienFromID(arNhanVienID(h))%></b>:				</td>
				<td align="left">&nbsp;&nbsp;<%=Dis_str_money(arNhanVienValues(h))&Donvigia%>				</td>
			</tr>
<%
		end if
	next
%>			
		</table>		</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent">Tổng tiền chi:</td>
        <td align="left" class="CTxtContent" valign="top" style="<%=setStyleBorder(0,0,0,1)%>"><b><%=Dis_str_money(fTongTienXuat)%></b></td>
        <td align="left" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent">Tổng tiền dư:</td>
        <td align="left" class="CTxtContent"><%=Dis_str_money(fTongTienThu-fTongTienXuat)%></td>
        <td align="right" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent"><strong>Ghi bằng chữ:</strong></td>
        <td colspan="2" align="left" class="CTxtContent"><strong><i><%=tienchu(fTongTienThu-fTongTienXuat)%></i></strong></td>
      </tr>
      <tr>
        <td colspan="3" align="left" class="CTxtContent">		</td>
      </tr>
</table>
<br>
<div align="center" class="author">
  <div align="left"><strong>NHẬP KHO</strong> </div>
</div>
<%if ctNhapKho = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	<tr>
	  <td width="7%"align="center" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
	  <td width="35%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Số</td>
	  <td width="6%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
	  <td width="14%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Mua hàng </td>
	  <td width="12%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Kiểm soát </td>
	  <td width="13%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Thanh Toán </td>
	  <td width="2%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"> SL </td>
	  <td width="11%"align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">T.Tiền</td>
	</tr>
<%end if%>	
    <%
	sql ="SELECT  inProductID,Maso,ProviderName,Ho_Ten,WorkerThanhToanID,AccountingID,DateTime FROM  inputProduct INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID INNER JOIN Nhanvien ON inputProduct.WorkerMuaHangID = Nhanvien.NhanVienID "
	sql = sql + "where (inputProduct.AccountingSigna<>0 or inputProduct.StoreSigna<>0 or inputProduct.CreaterSigna<>0) and (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0) "
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY DateTime desc"
	else
		sql=sql & " ORDER BY DateTime" 
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1 
	stt =1
	fTienNhap = 0
	Do while not rs.eof
		for h = 0 to iCountNV
			if arNhanVienID(h) = rs("WorkerThanhToanID") then
				arNhanVienChi(h) = arNhanVienChi(h) + LamTronTien(GetTTien(rs("inProductID")))
			end if
		next
		fTienNhap = fTienNhap + GetTTien(rs("inProductID"))
		tTienNhap = LamTronTien(GetTTien(rs("inProductID")))
%>
<%if ctNhapKho = 1 then%>		
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="Report_SoHD.asp?inProductID=<%=rs("inProductID")%>" target="_parent" class="CSubMenu"><%=rs("Maso")%></a><br>
		  <font class="CSubTitle">
		  <i>Nhà cung cấp</i>: <%=rs("ProviderName")%>		  </font>		  </td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("DateTime"))%>/<%=Month(rs("DateTime"))%>/<%=Year(rs("DateTime"))%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ho_Ten")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=getNhanVienFromID(rs("AccountingID"))%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%=getNhanVienFromID(rs("WorkerThanhToanID"))%>
		  </td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetTotalSPinHD(rs("inProductID"))%></td>
		  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(tTienNhap)%></td>
		</tr>
<%end if%>
<%
		stt = stt+1
		rs.movenext
	loop
%>
<%if ctNhapKho = 1 then%>
</table>
<%end if%>

   <br>
   
<table width="700" border="0" cellpadding="2" cellspacing="2">
      <tr>
        <td colspan="3" class="CTxtContent" >Tổng số lượng: <%=stt%> đơn hàng        </td>
      </tr>		  
      <tr>
        <td width="163" align="left" class="CTxtContent">Tổng chi:</td>
        <td width="140" align="left" valign="top" class="CTxtContent" style="<%=setStyleBorder(0,1,0,0)%>"><b><%=Dis_str_money(fTienNhap)%></b>		</td>
        <td width="377" align="left" class="CTxtContent">
		<table border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
<%
	for h = 0 to iCountNV-1 
	 if arNhanVienChi(h) <> "" and  arNhanVienChi(h) <> 0 then
%>			<tr>
				<td align="right">
				<b><%=getNhanVienFromID(arNhanVienID(h))%></b>:				</td>
				<td align="left">&nbsp;&nbsp;<%=Dis_str_money(arNhanVienChi(h))&Donvigia%>				</td>
			</tr>
<%
		end if
	next
%>			
		</table>		</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent"><strong>Ghi bằng chữ:</strong></td>
        <td colspan="2" align="left" class="CTxtContent"><strong><i><%=tienchu(fTienNhap)%></i></strong></td>
      </tr>
      <tr>
        <td colspan="3" align="left" class="CTxtContent">		</td>
      </tr>
</table> 
<br>  
 <div align="center" class="author">
  <div align="left"><strong>TỒN KHO</strong> </div>
</div>  
<%		
		strProd="SELECT inputProduct.inProductID, inputProduct.Maso, Product.ProductID, Product.NewsID,SanPhamNhap.Title,"
		strProd=strProd + "Provider.ProviderName, inputProduct.DateTime, Product.Number, Product.Giabia, Product.Price, Nhanvien.Ho_Ten,inputProduct.WorkerMuaHangID"
		strProd=strProd + " FROM Product Product INNER JOIN inputProduct ON Product.inProductID = inputProduct.inProductID"
		strProd=strProd + " INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID  "
		strProd=strProd + " INNER JOIN  Nhanvien ON inputProduct.AccountingID = Nhanvien.NhanVienID  "
		strProd=strProd + " INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID"					  
		strProd=strProd&" WHERE(DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) and (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0)"
		strProd=strProd + " and (inputProduct.AccountingSigna<>0 or inputProduct.StoreSigna<>0 or inputProduct.CreaterSigna<>0) "	
		if iOrderBy = 1 then 
			sql=sql & " ORDER BY DateTime desc"
		else
			sql=sql & " ORDER BY DateTime" 
		end if		
		dim rsProd
		set rsProd=Server.CreateObject("ADODB.Recordset")
		rsProd.open strProd,Con,1
		iSTT=1		 
	%>
<%if ctTonKho = 1 then%>	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	<tr align="center" valign="middle">
		<td width="3%" style="<%=setStyleBorder(1,1,1,1)%>"><strong>TT</strong></td>
		<td width="15%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã HĐ </strong></td>
		<td width="44%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên sách </strong></td>
		<td width="8%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Giá bán </strong></td>
		<td width="8%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL Nhập </strong></td>
		<td width="8%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL Xuất </strong></td>
		<td width="6%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL Trả </strong></td>
		<td width="8%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL Tồn</strong></td>
  </tr>
<%end if%>	  
		<%
		iTotalNhap =0
		iTotalXuat = 0
		fTongTon	= 0
		Do while not rsProd.eof
			SLxuat 	= 	GetNumInvoiceOutStore(rsProd("ProductID"))
			SL		=	GetNumeric(rsProd("Number"),0)
			kt		= 	SL - SLXuat - GetNumInvoiceReturnProvice(rsProd("ProductID"))
		if (kt > 0 or GetNumInvoiceReturnProvice(rsProd("ProductID")) > 0)then
			SLTonKho 	= 	kt
			fTongTon 	= 	fTongTon	+	rsProd("Price")*SLTonKho
			iTotalNhap 	= 	iTotalNhap	+ 	SL	
			iTotalXuat 	=	iTotalXuat	+ 	SLxuat	
		%>
<%if ctTonKho = 1 then%>		
		<tr  valign="middle">
		  <td align="center" width="3%" style="<%=setStyleBorder(1,1,0,1)%>"><%=iSTT%></td>
		  <td  align="Left" width="15%" style="<%=setStyleBorder(0,1,0,1)%>">
		 <a href="Report_SoHD.asp?inProductID=<%=rsProd("inProductID")%>" target="_parent" class="CSubMenu"><%=rsProd("Maso")%></a><br>
		 <font class="CSubTitle">
		  <i>Ngày nhập</i>: <%=Day(rsProd("DateTime"))%>/<%=Month(rsProd("DateTime"))%>/<%=Year(rsProd("DateTime"))%>
		  </font>
		  </td>
			<td  align="Left" width="44%" style="<%=setStyleBorder(0,1,0,1)%>">
			<i>Tiêu đề</i>: <%=rsProd("Title")%><br>
			<i>Giá bìa</i>: <%=Dis_str_money(rsProd("Giabia"))%><br>
			</td>
		    <td  align="right" width="8%" style="<%=setStyleBorder(0,1,0,1)%>">
			<%=Dis_str_money(rsProd("Price"))%></td>
			<td  align="center" width="8%" style="<%=setStyleBorder(0,1,0,1)%>"><%=SL%></td>
			<td  align="center" width="8%" style="<%=setStyleBorder(0,1,0,1)%>">
			<%=SLxuat%></td>
		    <td  align="center" width="6%" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetNumInvoiceReturnProvice(rsProd("ProductID"))%></td>
	      <td  align="Center" width="8%" style="<%=setStyleBorder(0,1,0,1)%>"><%=SLTonKho%></td>
	    </tr>
<%end if%>		
		<%
		iSTT=iSTT+1
		end if
		rsProd.movenext
		Loop
		
		%>
<%if ctTonKho = 1 then%>
</table>
<%end if%>
<br>
<table width="414" border="0" cellpadding="1" cellspacing="1">
  <tr>
    <td width="110"  class="CTxtContent">Tổng tồn kho:</td>
    <td width="151"  class="CTxtContent" style="<%=setStyleBorder(0,0,0,1)%>"><b><%=iTotalNhap-iTotalXuat%></b></td>
    <td width="143"  class="CTxtContent">&nbsp;</td>
  </tr>
  <tr>
    <td  class="CTxtContent">Tổng tiền tồn:</td>
    <td  class="CTxtContent"><b><%=Dis_str_money(fTongTon)&Donvigia%> </b></td>
    <td  class="CTxtContent">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3"  class="CTxtContent">Ghi bằng chữ:<b><%=tienchu(fTongTon)%></b></td>
  </tr>
</table>
<br>
 <div align="center" class="author">
  <div align="left"><strong>SÁCH XUẤT  TRONG KHO</strong> </div>
</div> 
<%if ctxseo = 1 then%>
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
   <tr align="center" valign="middle">
     <td width="4%" style="<%=setStyleBorder(1,1,1,1)%>"><strong>TT</strong></td>
     <td width="6%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã DH </strong></td>
     <td width="12%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã HD </strong></td>
     <td width="51%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên sách </strong></td>
     <td width="7%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Giá bìa </strong></td>
     <td width="7%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Giá nhập</strong></td>
     <td width="5%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL Xuất </strong></td>
     <td width="8%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thành tiền</strong> </td>
   </tr>
<%end if%> 
<%
		strProd="SELECT inputProduct.inProductID, inputProduct.Maso, Product.ProductID, Product.NewsID, SanPhamNhap.Title, inputProduct.DateTime, Product.Giabia, Product.Price, inputProduct.WorkerMuaHangID, XuatKho.SoLuong, inputProduct.DateTime, Product.Giabia, Product.Price, inputProduct.WorkerMuaHangID, XuatKho.SoLuong,SanPhamUser.NgayXuLy,SanPham_User.SanPhamUser_ID FROM Product INNER JOIN inputProduct ON Product.inProductID = inputProduct.inProductID INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID INNER JOIN XuatKho ON Product.ProductID = XuatKho.ProductID INNER JOIN SanPham_User ON XuatKho.SanPham_User_ID = SanPham_User.SanPham_User_ID INNER JOIN SanPhamUser ON SanPham_User.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID "
		strProd=strProd&" WHERE(DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0)"
		strProd=strProd&" AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
		strProd=strProd&" AND (DATEDIFF(dd, DateTime, '" & FromDate & "') > 0)"
		strProd=strProd + "  and inputProduct.AccountingSigna<>0 and inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0 "	
		if iOrderBy = 1 then 
			sql=sql & " ORDER BY NgayXuLy desc"
		else
			sql=sql & " ORDER BY NgayXuLy" 
		end if
		set rsProd=Server.CreateObject("ADODB.Recordset")
		rsProd.open strProd,Con,1
		iSTT=1	
		iTotalXuat= 0
		fTongTienKho=0
		Do while not rsProd.eof
		SanPhamUser_ID	= rsProd("SanPhamUser_ID")
		SL 	=	rsProd("SoLuong")
		iTotalXuat = iTotalXuat+ SL
		fTienKho = rsProd("Price")*SL
		fTongTienKho=fTongTienKho+fTienKho

		%>
<%if ctxseo = 1 then%>		
   <tr  valign="middle">
    <td align="center" width="4%" style="<%=setStyleBorder(1,1,0,1)%>"><%=iSTT%></td>
    <td  align="Left" width="6%" style="<%=setStyleBorder(0,1,0,1)%>">
	<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
      <%
			munb=1000+SanPhamUser_ID
			strTemp	="XB"+CStr(munb)
			Response.Write(strTemp)
		%></a>
	 </td>
     <td  align="Left" width="12%" style="<%=setStyleBorder(0,1,0,1)%>">
     <a href="Report_SoHD.asp?inProductID=<%=rsProd("inProductID")%>" target="_parent" class="CSubMenu"><%=rsProd("Maso")%></a></td>
     <td  align="Left" width="51%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProd("Title")%>     </td>
     <td  align="right" width="7%" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsProd("Giabia"))%></td>
     <td  align="right" width="7%" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsProd("Price"))%></td>
     <td  align="center" width="5%" style="<%=setStyleBorder(0,1,0,1)%>"><%=SL%></td>
     <td  align="center" width="8%" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(fTienKho)%></td>
   </tr>
   <%
		iSTT=iSTT+1
end if		
		rsProd.movenext
		Loop
		
		%>
<%if ctxseo = 1 then%>
</table>
<%end if%>
 <br>
 <table width="513" border="0" cellpadding="1" cellspacing="1">

   <tr>
     <td width="103"  class="CTxtContent">Tổng  xuất:</td>
     <td width="180"  class="CTxtContent" style="<%=setStyleBorder(0,0,0,1)%>"><b><%=iTotalXuat%></b></td>
     <td width="220"  class="CTxtContent">&nbsp;</td>
   </tr>

   <tr>
     <td  class="CTxtContent">Tổng tiền xuất:</td>
     <td  class="CTxtContent"><b><%=Dis_str_money(fTongTienKho)&Donvigia%> </b></td>
     <td  class="CTxtContent">&nbsp;</td>
   </tr>
   <tr>
     <td colspan="3" class="CTxtContent">Ghi bằng chữ:<b><%=tienchu(fTongTienKho)%></b></td>
   </tr>
</table>
<br>
<div align="center" class="author">
  <div align="left"><strong>ĐƠN HÀNG HỦY BỎ</strong> </div>
</div>
<font class="CSubTitle">Chịu phí bưu điện và chi phí khác</font>
<%
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM SanPhamUser " 
	sql=sql + " where SanPhamUser_Status= 3 "
	sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
	sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY NgayXuLy desc"
	else
		sql=sql & " ORDER BY NgayXuLy" 
	end if
	rs.open sql,con,3
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	end if
%>
<%if ctHuyBo = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr> 
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td width="54%" align="center" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tên</b></td>
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Giao hàng </td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Kiểm soát</b> </div></td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Chi </td>
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Ngày đặt</b> </td>
  </tr>
<%end if%>  
<%
iMau=0
STT = 0
fTongChiDHHuy =0
Do while not rs.eof 
SanPhamUser_ID		=	rs("SanPhamUser_ID")
SanPhamUser_Name	=	rs("SanPhamUser_Name")
SanPhamUser_Email	=	rs("SanPhamUser_Email")
SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
SanPhamUser_Address	=	rs("SanPhamUser_Address")
NgayXuLy	=	rs("NgayXuLy")
strCMND				=	rs("CMND")
KSoat				=	getNhanVienFromID(rs("KiemSoat"))
NVGiaoHang			=	getNhanVienFromID(rs("NhanVienID"))	
if KSoat = "" then
	KSoat="&nbsp;"
end if
if NVGiaoHang = "" then
	NVGiaoHang="&nbsp;"
end if
fChiDHHuy = GetCuocBuuDienThucID(SanPhamUser_ID) + GetChiKhac(SanPhamUser_ID)
fChiDHHuy = LamTronTien(fChiDHHuy)
fTongChiDHHuy = fTongChiDHHuy +	fChiDHHuy
%>
<%if ctHuyBo = 1 and fChiDHHuy > 0 then%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td width="7%"align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu"><%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a></td>
    <td align="left" valign="middle" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%><br>
	<font class="CSubTitle">	
	<i>Điện thoại</i>: <%=SanPhamUser_Tell%><br>
	<i>Địa chỉ</i>: <%=SanPhamUser_Address%></font>	</td>
    <td width="14%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=NVGiaoHang%></td>
	<td width="13%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%></td>
	<td width="6%"  style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(fChiDHHuy)%></td>
	<td width="6%" style="<%=setStyleBorder(0,1,0,1)%>">
		<%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%>	</td>
  </tr>

<%
iMau=iMau+1
end if
stt=stt + 1
rs.movenext
Loop
%>
<%if ctHuyBo = 1 then%>
</table>
<%end if%>
<%	rs.close
	set rs=nothing
%>	
<br>
    <table width="700" border="0" cellpadding="2" cellspacing="2">
      <tr>
        <td width="161" align="left" class="CTxtContent">Tổng tiền chi:</td>
        <td width="141" align="left" class="CTxtContent"><%=Dis_str_money(fTongChiDHHuy)%></td>
        <td width="378" align="right" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent"><strong>Ghi bằng chữ:</strong></td>
        <td colspan="2" align="left" class="CTxtContent"><strong><i><%=tienchu(fTongChiDHHuy)%></i></strong></td>
      </tr>
      <tr>
        <td colspan="3" align="left" class="CTxtContent">		</td>
      </tr>
</table>
<br>
 <div align="center" class="author">
  <div align="center"><strong>TỔNG KẾT</strong> </div>
</div> 
 <table width="500" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFCC">
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,1,1)%>">Tổng tiền thu: </td>
     <td align="right" class="CTxtContent" style="<%=setStyleBorder(1,1,1,1)%>"><%=Dis_str_money(fTongTienThu)&Donvigia%></td>
   </tr>
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,0,1)%>">Tổng tiền chi: </td>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,1,0,1)%>">- <%=Dis_str_money(fTongTienXuat)&Donvigia%></td>
   </tr>
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,0,1)%>">Tổng tiền tồn kho: </td>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,1,0,1)%>">- <%=Dis_str_money(fTongTon)&Donvigia%></td>
   </tr>
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,0,1)%>">Tổng tiền sách xuất  kho xseo: </td>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,1,0,1)%>">
	 + <%=Dis_str_money(fTongTienKho)&Donvigia%> </td>
   </tr>
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,0,1)%>">Tổng chi đơn hàng hủy: </td>
     <td  class="CTxtContent" align="right" style="border-bottom:#FF0000 solid 2;border-left:#333333 solid 1px; border-right:#333333 solid 1">
	 - <%=Dis_str_money(fTongChiDHHuy)&Donvigia%></td>
   </tr>
   <tr>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,0,0,1)%>">Số dư còn lại:</td>
     <td  class="CTxtContent" align="right" style="<%=setStyleBorder(1,1,0,1)%>">
       <%fDuConLai	=	fTongTienThu - fTongTienXuat - fTongTon + fTongTienKho - fTongChiDHHuy%>
     <%=Dis_str_money(fDuConLai)&Donvigia%> </td>
   </tr>
   <tr>
     <td colspan="2"  class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>">Ghi bằng chữ:<b><%=tienchu(fDuConLai)%></b></td>
   </tr>
</table>

<br>  
<p>&nbsp;</p>
</body>
</html>
