<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%
	Maso			=	trim(Request.Form("txtMaso"))
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
	NhCC=Clng(Request.Form("selProvider"))
	WorkerKSID=GetNumeric(Request.Form("selKS"),0)
	WorkerMHID=GetNumeric(Request.Form("selMH"),0)
	WorkerTTID=GetNumeric(Request.Form("selTT"),0)
	FromDate=Thang1&"/"&Ngay1&"/"&Nam1
	ToDate=Thang2&"/"&Ngay2&"/"&Nam2
		
%>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<SCRIPT language=JavaScript1.2 src="../administrator/inc/calendarDateInput.js"></SCRIPT>
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
  <div align="center"> <font class="CTieuDe">THỐNG KÊ NHẬP HÀNG</font><br>
  <font class="CTxtContent"><%=Day(now)%>/<%=Month(Now)%>/<%=year(Now)%></font>
  </div>

<br>
<%
	sqlnv	= "Select ProviderID From Provider"
	set rsncc = Server.CreateObject("ADODB.recordset")
	rsncc.open sqlnv,con,1
	iCount	= rsncc.recordcount - 1
	Redim arNCC(iCount)
	redim arNCCValues(iCount)
	h = 0
	do while not rsncc.eof
		arNCC(h) = rsncc("ProviderID")
		h= h +1
		rsncc.movenext
	loop
	set rsncc = nothing	
	iTotal =0
	iTotalSL = 0
	iDonGia = 0
if iDetail = 1 then
		strProd="SELECT inputProduct.inProductID,inputProduct.Maso,Product.NewsID,SanPhamNhap.Title,Provider.ProviderID,Provider.ProviderName,inputProduct.DateTime"
		strProd=strProd + ",Product.Unit,Product.Number,Product.Giabia,Product.Price,Product.VAT,Nhanvien.Ho_Ten"
		strProd=strProd + ",inputProduct.WorkerMuaHangID"
		strProd=strProd + " FROM Product INNER JOIN inputProduct ON Product.inProductID = inputProduct.inProductID "
		strProd=strProd + " INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID "
		strProd=strProd + " INNER JOIN Nhanvien ON inputProduct.AccountingID = Nhanvien.NhanVienID "
		strProd=strProd + " INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID"					  
		strProd=strProd&" WHERE  inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0 and AccountingSigna<>0 "
		strProd=strProd&" and (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0)"

		If (NhCC <> "0") Then 
			strProd = strProd&" AND inputProduct.ProviderID="&NhCC
		End If
		If (WorkerKSID <> "0") Then 
			strProd = strProd&" AND inputProduct.AccountingID="&WorkerKSID
		End If
		If (WorkerMHID <> "0") Then 
			strProd = strProd&" AND inputProduct.WorkerMuaHangID="&WorkerMHID
		End If
		If (WorkerTTID <> "0") Then 
			strProd = strProd&" AND inputProduct.WorkerThanhToanID="&WorkerTTID
		End If

		select case iMaorTenSach 
			case 1
			strProd = strProd + " and inputProduct.Maso = '"& strMaorTenSach &"'  "
			case 2
			strProd = strProd + " and SanPhamNhap.Title like N'%"& strMaorTenSach &"%'"
		end select
		
		if iOrderBy = 1 then 
			strProd=strProd & " ORDER BY "& strSelSearch &" desc"
		else
			strProd=strProd & " ORDER BY "& strSelSearch 
		end if

		set rsProd=Server.CreateObject("ADODB.Recordset")
		rsProd.open strProd,Con,1
		iSTT=1		 
	%>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td colspan="2" align="left" class="CTxtContent">Thời gian:<b> Từ:<%=GetFullDate(FromDate)%>&nbsp;đến:<%=GetFullDate(ToDate)%></b></td>
      </tr>
	  	<% 
			if iMaorTenSach = 1 then
			sql11 = "SELECT  Maso, ProviderID,WorkerMuaHangID, AccountingID,WorkerThanhToanID, DateTime FROM  inputProduct"
			sql11 = sql11 + " where Maso = '"& strMaorTenSach &"'"
			if NhCC <>0 then
				sql11 = sql11 + " and ProviderID = '"& NhCC &"'"
			end if
			set rsTemp = Server.CreateObject("ADODB.recordset")
			rsTemp.open sql11,Con,1
			do while not rsTemp.eof
				WorkerKSID	=	GetNumeric(rsTemp("AccountingID"),0)
				WorkerMHID	=	GetNumeric(rsTemp("WorkerMuaHangID"),0)
				WorkerTTID	=	GetNumeric(rsTemp("WorkerThanhToanID"),0)
				if NhCC = 0 then
					NhCC		=	GetNumeric(rsTemp("ProviderID"),0)
				end if
				DateDonHang	=  	rsTemp("DateTime")			
				rsTemp.movenext
			loop
			rsTemp.close
			set rsTemp = nothing
		
		%>

      <tr>
        <td align="left" class="CTxtContent">Mã HD:<b><%=strMaorTenSach%></b></td>
        <td width="640" align="left" class="CTxtContent">Ngày:<%=Day(DateDonHang)%>/<%=Month(DateDonHang)%>/<%=year(DateDonHang)%></td>
      </tr>
	  	 <%
	  end if
	  %>
		<%if NhCC <> 0 then %>
		<tr>
        <td align="left" class="CTxtContent">
			Nhà cung cấp:<b><%=getProviderFormID(NhCC)%></b>
		</td>
        <td align="left" class="CTxtContent">&nbsp;</td>
      </tr>
<%
		end if
		if WorkerMHID <> 0 then
%>
      <tr>
        <td align="left" class="CTxtContent">
			Người mua:<b><%=getNhanVienFromID(WorkerMHID)%></b>
		</td>
        <td align="left" class="CTxtContent">
			
		</td>
      </tr>
<%		end if
		if WorkerKSID <> 0 then%>
      <tr>
        <td width="273" align="left" class="CTxtContent">
		Kiểm soát viên:<b><%=getNhanVienFromID(WorkerKSID)%><b>
		</td>
        <td align="left" class="CTxtContent">		</td>
      </tr>
<%	  		end if
if WorkerTTID <> 0 then
		%>	
      <tr>
        <td align="left" class="CTxtContent">
		Thanh toán:<b><%=getNhanVienFromID(WorkerTTID)%></b>

		</td>
        <td align="left" class="CTxtContent"></td>
      </tr>
<%end if%>
</table>
	<br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000" class="CTxtContent">
		<tr align="center" valign="middle" bgcolor="#FFFFFF">
			<td width="2%" style="<%=setStyleBorder(1,1,1,1)%>"><strong>TT</strong></td>
			<% if iMaorTenSach <> 1 then%>
				<td width="4%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã HĐ </strong></td>
			<%
			end if
			
			if NhCC = 0 and iMaorTenSach <> 1 then
			%>	
		    <td width="6%" style="<%=setStyleBorder(0,1,1,1)%>">Nhà CC </td>
			<%
				end if
			 if iMaorTenSach <> 1 then
			%>
		        <td width="4%" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
			<%
			end if
			if WorkerKSID = 0  and iMaorTenSach <> 1then
			%>
            <td width="12%" style="<%=setStyleBorder(0,1,1,1)%>">Kiểm SV </td>
			<%end if
			if WorkerMHID = 0  and iMaorTenSach <> 1 then
			%>
			<td width="13%" style="<%=setStyleBorder(0,1,1,1)%>">Mua hàng </td>
			<%end if%>
            <td width="40%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên hàng</strong></td>
			<td width="2%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL</strong></td>
			<td width="4%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Đơn giá 
			</strong></td>
			<td width="4%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>VAT
			</strong></td>
			<td width="9%" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thành tiền</strong></td>
	  </tr>  
		<%
		Do while not rsProd.eof
			Price 	= 	GetNumeric(rsProd("Price"),0)
			SL		=	GetNumeric(rsProd("Number"),0)	
			iVAT  =  rsProd("VAT")
			TTien = Price*SL
			iDonGia  = iDonGia  + TTien
			iTotalSL = iTotalSL + SL
			TTien = TTien + TTien*iVAT/100	
			iTotal = iTotal + TTien
			for h = 0 to ubound(arNCC) 
				if arNCC(h) = rsProd("ProviderID") then
					arNCCValues(h)	= 	arNCCValues(h) + TTien
				end if
			next			
			%>
		<tr  valign="middle" bgcolor="#FFFFFF">
		  <td align="center" width="2%" style="<%=setStyleBorder(1,1,0,1)%>"><%=iSTT%></td>
		  <% if iMaorTenSach <> 1 then%>
		  <td  align="Left" width="4%" style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="Report_SoHD.asp?inProductID=<%=rsProd("inProductID")%>" target="_parent" class="CSubMenu">
		 	 <%=rsProd("Maso")%>		  </a>		  </td>
			<%
			end if
				if NhCC = 0 and iMaorTenSach <> 1 then
			%>	
			<td  align="Left" width="6%" style="<%=setStyleBorder(0,1,0,1)%>"> <%=rsProd("ProviderName")%></td>
			<%end if
			if iMaorTenSach <> 1 then
			%>
			<td  align="Left" width="4%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProd("DateTime")%></td>
			<%
			end if
			if WorkerKSID = 0  and iMaorTenSach <> 1 then%>
		  <td  align="Left" width="12%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProd("Ho_Ten")%></td>
		  <%end if
		  	if WorkerMHID = 0  and iMaorTenSach <> 1 then
		  %>
			<td  align="Left" width="13%" style="<%=setStyleBorder(0,1,0,1)%>"><%=getNhanVienFromID(rsProd("WorkerMuaHangID"))%></td>
			<%end if%>
			<td  align="Left" width="40%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProd("Title")%></td>
			<td  align="center" width="2%" style="<%=setStyleBorder(0,1,0,1)%>"><%=SL%></td>
			<td  align="Right" width="4%" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(Price)%></td>
			<td  align="Center" width="4%" style="<%=setStyleBorder(0,1,0,1)%>"><%=iVAT%></td>

			<td  align="Right" width="9%" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(TTien)%></td>
	    </tr>

		<%
		iSTT=iSTT+1
		rsProd.movenext
		Loop
		%>
		
</table>
<br>
<br>
<%else%>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
		<tr>
		  <td width="5%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(1,1,1,1)%>" height="25">STT</td>
		  <td width="11%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Số</td>
		  <td width="11%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
		  <td width="18%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Nhà cung cấp</td>
		  <td width="14%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Kiểm soát </td>
		  <td width="14%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Thanh Toán </td>
		  <td width="7%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tổng SL </td>
		  <td width="9%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">T.Tiền</td>
		  <%if Session("iQuanTri") = 1 then %>
		  <td width="11%"align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
		  <%end if%>
		</tr>
<%
	sql ="SELECT  inputProduct.inProductID,inputProduct.Maso,inputProduct.ProviderID,ProviderName,Ho_Ten,WorkerThanhToanID,DateTime"
	if iMaorTenSach = 2 then
		sql = sql + ",Product.Number "
	end if
	sql = sql + " FROM  inputProduct INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID INNER JOIN Nhanvien ON inputProduct.AccountingID = Nhanvien.NhanVienID "
	if iMaorTenSach <> 0 then
		sql = sql + " INNER JOIN Product ON inputProduct.inProductID = Product.inProductID INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID "
	end if
	sql = sql + " where inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0  and inputProduct.AccountingSigna<>0 and (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0) "
	
	If (NhCC <> "0") Then 
		sql = sql&" AND inputProduct.ProviderID="&NhCC
	End If
	If (WorkerKSID <> "0") Then 
		sql = sql&" AND inputProduct.AccountingID="&WorkerKSID
	End If
	If (WorkerMHID <> "0") Then 
		sql = sql&" AND inputProduct.WorkerMuaHangID="&WorkerMHID
	End If
	If (WorkerTTID <> "0") Then 
		sql = sql&" AND inputProduct.WorkerThanhToanID="&WorkerTTID
	End If
	select case iMaorTenSach 
		case  2
			sql = sql + " and (Title like N'%"& strMaorTenSach &"%')"
		case  1
			sql = sql + " and inputProduct.Maso = '"& strMaorTenSach &"'  "
	end select
	sql = sql + "  ORDER BY DateTime DESC "
'	Response.Write(sql)
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1 
	stt =1
	Do while not rs.eof
		TTien	=	LamTronTien(GetTTien(rs("inProductID")))
		iDonGia  = iDonGia  + TTien
		if iMaorTenSach = 2 then
			SL = rs("Number")
		else
			SL		=	GetTotalSPinHD(rs("inProductID"))
		end if
		iTotalSL = iTotalSL + SL	
		iTotal = iTotal + TTien
		for h = 0 to ubound(arNCC) 
			if arNCC(h) = rs("ProviderID") then
				arNCCValues(h)	= 	arNCCValues(h) + TTien
			end if
		next			
	%>
		<tr>
		  <td align="center" height="26" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="Report_SoHD.asp?inProductID=<%=rs("inProductID")%>" target="_parent" class="CSubMenu">
		  <%=rs("Maso")%>
		  </a>&nbsp;</td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("DateTime"))%>/<%=Month(rs("DateTime"))%>/<%=Year(rs("DateTime"))%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("ProviderName")%>&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ho_Ten")%>&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=getNhanVienFromID(rs("WorkerThanhToanID"))%>&nbsp;</td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=SL%></td>
		  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(TTien)%>
		  </td>
		  <%if Session("iQuanTri") = 1 then %>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
			<a href="../InputProduct/InputProductList.asp?inProductID=<%=rs("inProductID")%>">Sửa</a>
			|<a href="javascript:winpopup('../InputProduct/DelHD.asp','<%=rs("inProductID")%>',300,150);">Xóa</a>
			
		  </td><%end if%>
		</tr>
	<%
	stt = stt+1
	rs.movenext
	loop
	%>
	<tr>
	<td colspan="6" align="center" height="28"><b>Tổng</b></td>
	<td align="center"><%=Dis_str_money(iTotalSL)%>	</td>
	<td align="right"><b><%=Dis_str_money(iTotal)&DonviGia%></b>	</td>		
	<%if Session("iQuanTri") = 1 then %>
	<td align="center">&nbsp;
	</td>
	<%end if%>
	</tr>
</table>	
<%end if%>

<%if iBieuDo = 1 then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <th scope="row"><table border="0" cellspacing="0" cellpadding="0" class="CTxtContent">
      <tr>
        <td valign="top" align="right"> Triệu </td>
        <td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
        <%for i = 0 to ubound(arNCC)
	if	arNCCValues(i) > 10000 then
	%>
        <td width="50" align="center" valign="bottom"><%=round(arNCCValues(i)/1000000,3)%><br>
            <%for k = 0 to round(arNCCValues(i)/100000)
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
	next
	end if
	%>        </td>
        <%next%>
        <td width="25" align="center" valign="bottom"></td>
      </tr>
      <tr>
        <td align="right" valign="top">&nbsp;</td>
        <td valign="top"><br></td>
        <%for i = 0 to ubound(arNCC)
		if	arNCCValues(i) > 10000 then
	%>
        <td style="border-top:#000000 solid 1" align="center" valign="top"><%=getProviderFormID(arNCC(i))%> </td>
        <%
		end if
	next
		
		%>
        <td style="border-top:#000000 solid 1" align="right" valign="top" width="50"> NCC </td>
      </tr>
    </table></th>
  </tr>
  <tr>
    <th scope="row"class="CFontVerdana10"><br>
    Tổng tiền nhập của từng nhà cung cấp
    <br></th>
  </tr>
</table>
<%end if%>
<br>
<%if iDetail = 1 then%>
<table width="700" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFCC" style="border:#CCCCCC solid 1">
      <tr>
        <td width="105" align="right" class="CTxtContent">Tổng số lượng: </td>
        <td width="75" align="left" class="CTxtContent"><%=Dis_str_money(iTotalSL)%> cuốn</td>
        <td width="189" align="right" class="CTxtContent">Giá trước thuế:</td>
        <td width="50" align="left" class="CTxtContent"><b><%=Dis_str_money(iDonGia)&DonviGia%></b></td>
        <td width="183" align="right" class="CTxtContent">Tổng sau thuế:</td>
        <td width="79" align="left" class="CTxtContent"><b><%=Dis_str_money(iTotal)&DonviGia%></b></td>
      </tr>
      <tr>
        <td colspan="6" align="right"class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="6" align="center" class="CTxtContent">Ghi bằng chữ: <b><i><%=tienchu(iTotal)%></i></b></td>
      </tr>
</table>
<%end if%>
<br>    
</body>
</html>
