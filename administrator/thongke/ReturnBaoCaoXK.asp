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
if f_permission = 0 then
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
iDetail			=	GetNumeric(Request.form("iDetail"),0)
iBieuDo			=	GetNumeric(Request.form("iBieuDo"),0)
strSelSearch	=	Trim(Request.Form("selSearch"))
iOrderBy		=  	Clng(Request.Form("RaOderBy"))
iMaorTenSach	=	Clng(Request.Form("selMaorTenSach"))
strMaorTenSach	=	Trim(Request.Form("txtMaOrTensach"))
WorkerKSID		=	GetNumeric(Request.Form("selKS"),0)
WorkerMHID		=	GetNumeric(Request.Form("selMH"),0)
WorkerThuTienID	=	GetNumeric(Request.Form("selNVThutien"),0)



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
    <em>ĐT: <%=soDT%>  - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%></em></div></td>
  </tr>
</table>
<br><br>
  <div align="center"class="CTieuDe">
    THỐNG KÊ XUẤT HÀNG
</div>
  <center> <%=Day(now)%>/<%=Month(Now)%>/<%=year(Now)%></center>
<br>

<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	Set rs=Server.CreateObject("ADODB.Recordset") 
	sql="SELECT * FROM SanPhamUser " 
	sql=sql + " where SanPhamUser_Status= 2 "	
	If (WorkerKSID <> "0") Then 
		sql = sql&" AND KiemSoat="&WorkerKSID
	End If
	If (WorkerMHID <> "0") Then 
		sql = sql&" AND NhanVienID="&WorkerMHID
	End If
	If (WorkerThutienID <> "0") Then 
		sql = sql&" AND NVThutienID="&WorkerThuTienID
	End If
	select case iMaorTenSach 
		case 1
			strMaorTenSach = 	replace(strMaorTenSach,"XB","")			
			if isnumeric(strMaorTenSach) = true then
				numb = Clng(strMaorTenSach) - 1000
			else
				numb = 0
			end if	
			sql = sql + " and SanPhamUser.SanPhamUser_ID = "&numb
		case 3
			sql = sql + " and {fn UCASE(SanPhamUser_Name)} like N'%"& UCase(strMaorTenSach) &"%'"
		case 4
			sql = sql + " and SanPhamUser_Email like N'%"& strMaorTenSach &"%'"
		case 5
			sql = sql + " and SanPhamUser_Tell like N'%"& strMaorTenSach &"%'"			
		case 6
			sql="SELECT  SanPhamUser.*,SanPhamNhap.idsanpham,SanPhamNhap.Title FROM SanPhamUser INNER JOIN SanPham_User ON SanPhamUser.SanPhamUser_ID = SanPham_User.SanPhamUser_ID INNER JOIN SanPhamNhap ON SanPham_User.SanPham_ID = SanPhamNhap.NewsID " 
			sql=sql + " where SanPhamUser_Status= 2 "		
			sql = sql + " and SanPhamNhap.idsanpham like N'%"& strMaorTenSach &"%'"			
		case 8
			sql="SELECT  SanPhamUser.*,SanPhamNhap.idsanpham,SanPhamNhap.Title FROM SanPhamUser INNER JOIN SanPham_User ON SanPhamUser.SanPhamUser_ID = SanPham_User.SanPhamUser_ID INNER JOIN SanPhamNhap ON SanPham_User.SanPham_ID = SanPhamNhap.NewsID " 
			sql=sql + " where SanPhamUser_Status= 2 "				
			sql = sql + " and SanPhamNhap.Title like N'%"& strMaorTenSach &"%'"			

		case 7
			sql = sql + " and {fn UCASE(GiaoHang_Address)} like N'%"& UCase(strMaorTenSach) &"%'"	
					
	end select
	sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
	sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY "& strSelSearch &" desc"
	else
		sql=sql & " ORDER BY "& strSelSearch 
	end if
	rs.open sql,con,3
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
		Response.End()
	else
		if iMaorTenSach  = 1 then
			Response.Redirect("ReportXKChiTiet.asp?SanPhamUser_ID="&rs("SanPhamUser_ID"))
		end if	
	end if

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td width="125">Từ ngày:<strong> <%=Ngay1%>/<%=Thang1%>/<%=Nam1%></strong></td>
    <td>đến ngày:<strong><%=Ngay2%>/<%=Thang2%>/<%=Nam2%></strong></td>
  </tr>
 <%	select case iMaorTenSach 
		case 1
%>			<tr>
				<td width="125">Số hóa đơn: </td>
	<td width="846"><strong><%=Trim(Request.Form("txtMaOrTensach"))%></strong></td>
			</tr>
<%
		case 3
%>
  <tr>
	<td>Tên:</td>
				<td><strong><%=strMaorTenSach%></strong></td>
  </tr>
<%
		case 4
%>
		  <tr>
			<td>Email:</td>
			<td><strong><%=strMaorTenSach%></strong></td>
		  </tr>
<%		case 5
%>		  <tr>
	<td>Tel:</td>
			<td><strong><%=strMaorTenSach%></strong></td>
		  </tr>
<%
  	end select
%>			
<%	If (WorkerKSID <> "0") Then 
%>
  <tr>
	<td>Kiểm soát viên: </td>
		<td><strong><%=getNhanVienFromID(WorkerKSID)%></strong></td>
  </tr>
<%		
	End If
	If (WorkerMHID <> "0") Then 
%>
  <tr>
	<td>Nhân viên giao hàng: </td>
		<td><strong><%=getNhanVienFromID(WorkerMHID)%></strong></td>
  </tr>
<%
	End If
	if WorkerThuTienID <> "0" then
%>	
  <tr>
    <td>NV thu tiền : </td>
    <td><strong><%=getNhanVienFromID(WorkerThuTienID)%></strong></td>
  </tr>	
<%
	end if
%>
</table>
<br>
<%
	sqlnv	= "Select NhanVienID From NhanVien"
	set rsnv = Server.CreateObject("ADODB.recordset")
	rsnv.open sqlnv,con,1
	iCountNV	=rsnv.recordcount - 1
	Redim arNhanVienID(iCountNV)
	redim arNhanVienValues(iCountNV)
	redim arKiemSoatValues(iCountNV)
	redim arGiaoHangValues(iCountNV)	
	h = 0
	do while not rsnv.eof
		arNhanVienID(h) = rsnv("NhanVienID")
		h= h +1
		rsnv.movenext
	loop
	set rsnv = nothing	
%>
<%if iDetail = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr> 
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tên</b></td>
<%if iMaorTenSach <> 5 then%>    
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tel</b></td>
<%end if%>	
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Địa chỉ</b></td>
	<%If WorkerMHID = "0" then %>
    <td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Giao hàng </td>
	<%end if%>
    <%If (WorkerKSID = "0") then%>	
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Kiểm soát</b> </div></td>
<%end if%>	
    <%If (WorkerThutienID = "0") then%>	
	<td class="CTextStrong" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Thu tiền</b> </div></td>
<%end if%>	
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
for iKS = 0 to iCountNV
	if arNhanVienID(iKS)= rs("KiemSoat") then
		arKiemSoatValues(iKS) = arKiemSoatValues(iKS) + 1
	end if
next
NVGiaoHang			=	getNhanVienFromID(rs("NhanVienID"))	
for iKS = 0 to iCountNV
	if arNhanVienID(iKS)= rs("NhanVienID") then
		arGiaoHangValues(iKS) = arGiaoHangValues(iKS) + 1
	end if	
next
NVThutien			=	getNhanVienFromID(rs("NVThutienID"))	
if KSoat = "" then
	KSoat="&nbsp;"
end if
if NVGiaoHang = "" then
	NVGiaoHang="&nbsp;"
end if
iTiepTuc =  1
if iMaorTenSach = 6 then
	NewsID = GetNewsIDFromIDSanPhamNhap(strMaorTenSach)
	if isCheckSanPhamUserID(SanPhamUser_ID,NewsID) = false then
		iTiepTuc = 0
	end if	
end if
if iTiepTuc = 1 then
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
		fTongXuat =LamTronTien(fTongXuat + GetCuocBuuDienThucID(SanPhamUser_ID) + GetChiKhac(SanPhamUser_ID))
		fTongTienXuat = fTongTienXuat + fTongXuat 
		
		iTien 	= 	0
		iTien 	= 	LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND))
		fTongTienThu =	fTongTienThu + iTien
		for h = 0 to iCountNV
			if arNhanVienID(h) = rs("NVThutienID") then
				arNhanVienValues(h) = arNhanVienValues(h) + iTien
			end if
		next
%>
<%if iDetail = 1 then%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td width="3%"align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu"><%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a>
	</td>
    <td width="12%" valign="middle" align="left" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%></font></td>
<%if iMaorTenSach <> 5 then%> 
    <td width="7%"align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=SanPhamUser_Tell%></td>
<%end if%>	
    <td width="30%"align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=SanPhamUser_Address%></td>
	<%If WorkerMHID = "0" then %>
    <td width="12%" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=NVGiaoHang%></td>
	<%end if%>
    <%If (WorkerKSID = "0") then%>		<td width="12%" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%></td><%end if%>
    <%If (WorkerThutienID = "0") then%>		<td width="12%" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=NVThutien%></td><%end if%>	
	<td width="9%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(fTongXuat)%></td>
	<td width="10%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(iTien)%>	</td>
	<td width="8%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(iTien-fTongXuat)%>	</td>
	<td width="9%" style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%></div></td>
  </tr>
<%end if%>
<%
SoDH = SoDH+1
stt=stt + 1
iMau=iMau+1
end if ' sử dụng cho biến iTiepTuc
rs.movenext
Loop%>
<%if iDetail = 1 then%>
</table>
<%end if%>
<%	rs.close
	set rs=nothing
%>	
<br>
<%if iBieuDo = 1 then%>
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <th scope="row"><span class="CFontVerdana10">Kiểm xoát viên </span></th>
  </tr>
  <tr>
    <th scope="row"><table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
      <tr>
        <td width="18" valign="top">Số<br>
          ĐH</td>
        <td width="4" valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
        <%for i = 0 to iCountNV
			if arKiemSoatValues(i) > 0 then
			%>
        <td width="150" align="center" valign="bottom">
			<%=arKiemSoatValues(i)%><br>
            <%for k = 0 to arKiemSoatValues(i)
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
			next%>			</td>
        <%
			end if
		next%>
        <td width="50" align="center" valign="bottom"></td>
      </tr>
      <tr>
        <td align="right" valign="top">&nbsp;</td>
        <td valign="top"><br></td>
        <%for i = 0 to iCountNV
			if arKiemSoatValues(i) > 0 then
		%>
        <td style="border-top:#000000 solid 1" align="center" valign="top"><%=getNhanVienFromID(arNhanVienID(i))%> </td>
        <%
			end if
		next%>
        <td style="border-top:#000000 solid 1" align="right" valign="top" width="25"> <div align="right">Tên </div></td>
      </tr>
    </table>      </th>
  </tr>
  <tr>
    <th class="CFontVerdana10" scope="row"><p>&nbsp;</p>
      <p><br>    	
    Nhân viên giao hàng </p></th>
  </tr>
  <tr>
    <th scope="row"><table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
      <tr>
        <td valign="top">Số<br>
          ĐH</td>
        <td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
        <%for i = 0 to iCountNV
			if arGiaoHangValues(i) > 0 then
		%>
        <td width="150" align="center" valign="bottom"><%=arGiaoHangValues(i)%><br>
            <%for k = 0 to arGiaoHangValues(i)
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
		next%>        </td>
        <%
			end if
		next%>
        <td width="50" align="center" valign="bottom"></td>
      </tr>
      <tr>
        <td align="right" valign="top">&nbsp;</td>
        <td valign="top"><br></td>
        <%for i = 0 to iCountNV
			if arGiaoHangValues(i) > 0 then
		%>
        <td style="border-top:#000000 solid 1" align="center" valign="top"><%=getNhanVienFromID(arNhanVienID(i))%> </td>
        <%
			end if
		next%>
        <td style="border-top:#000000 solid 1" align="right" valign="top" width="50"> <div align="right">Tên </div></td>
      </tr>
    </table></th>
  </tr>
  <tr>
    <th class="CFontVerdana10" scope="row">&nbsp;</th>
  </tr>
</table>
<%end if%>
<br>
    <table width="700" border="0" align="center" cellpadding="2" cellspacing="2">
      <tr>
        <td colspan="4" class="CTxtContent" align="center">Tổng số lượng: <%=stt%> đơn hàng
        <div align="left"></div></td>
      </tr>		  
      <tr>
        <td width="19" align="left" class="CTxtContent">&nbsp;</td>
        <td width="173" align="right" valign="top" class="CTxtContent">Tổng tiền thu:</td>
        <td width="55" align="left" valign="top" class="CTxtContent" style="<%=setStyleBorder(0,1,0,0)%>"><b><%=Dis_str_money(fTongTienThu)%></b>		</td>
        <td width="227" align="left" class="CTxtContent">
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
        <td align="left" class="CTxtContent">&nbsp;</td>
        <td align="right" class="CTxtContent" valign="top">Tổng tiền chi:</td>
        <td align="left" class="CTxtContent" valign="top" style="<%=setStyleBorder(0,0,0,1)%>"><b><%=Dis_str_money(fTongTienXuat)%></b></td>
        <td align="left" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent">&nbsp;</td>
        <td align="right" class="CTxtContent">Tổng tiền dư:</td>
        <td align="left" class="CTxtContent"><%=Dis_str_money(fTongTienThu-fTongTienXuat)%></td>
        <td align="right" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td align="left" class="CTxtContent">&nbsp;</td>
        <td align="right" class="CTxtContent"><strong>Ghi bằng chữ:</strong></td>
        <td colspan="2" align="left" class="CTxtContent"><strong><i><%=tienchu(fTongTienThu-fTongTienXuat)%></i></strong></td>
      </tr>
	  
 
      <tr>
        <td colspan="4" align="left"class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="4" align="left" class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="4" align="left" class="CTxtContent">		</td>
      </tr>
</table>
<br>
  
</body>
</html>
<%
function isCheckSanPhamUserID(SanPhamUser_ID,NewsID)
	sql = "select SanPham_ID From SanPham_User where SanPhamUser_ID ='"& SanPhamUser_ID &"' and SanPham_ID ='"& NewsID &"'  and re_newsid = 0"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,Con,1	
	if not rsTemp.eof then
		isCheckSanPhamUserID = true
	else
		isCheckSanPhamUserID = false
	end if
	rsTemp.close
	set rsTemp = nothing
	
end function
%>