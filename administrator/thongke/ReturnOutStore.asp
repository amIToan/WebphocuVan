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
	actions 	=	Request.QueryString("action")
	Maso		=	Request.Form("txtMaso")
	IF actions="Search" then
		Ngay1			=	GetNumeric(Request.form("Ngay1"),0)
		Thang1			=	GetNumeric(Request.form("Thang1"),0)
		Nam1			=	GetNumeric(Request.form("Nam1"),0)
		Ngay2			=	GetNumeric(Request.form("Ngay2"),0)
		Thang2			=	GetNumeric(Request.form("Thang2"),0)
		Nam2			=	GetNumeric(Request.form("Nam2"),0)
		strSelSearch	=	Trim(Request.Form("selSearch"))
		iOrderBy		=  	Clng(Request.Form("RaOderBy"))
		iMaorTenSach	=	Clng(Request.Form("selMaorTenSach"))
		strMaorTenSach	=	Trim(Request.Form("txtMaOrTensach"))
		WorkerKSID=GetNumeric(Request.Form("selKS"),0)
		WorkerMHID=GetNumeric(Request.Form("selMH"),0)
	ELSE
		Ngay1=Day(now())
		Thang1=Month(now())-1
		Nam1=Year(now())
		Ngay2=Day(now())
		Thang2=Month(now())
		Nam2=Year(now())
	END IF

FromDate=Thang1&"/"&Ngay1&"/"&Nam1
ToDate=Thang2&"/"&Ngay2&"/"&Nam2
		
%>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css"></head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
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
<div align="center" class="author">
  <div align="center"><strong>THỐNG KÊ XUẤT KHO </strong></div>
</div>
<br>
<br>
<%IF Request.QueryString("action")="Search" then
		strProd="SELECT SanPhamUser_Name,SanPhamUser_Email,SanPham_ID,SanPham_Gia,SanPham_Soluong, News.Title,SanPhamUser_ID,NhanVienID,KiemSoat,SanPhamUser_Date"
		strProd=strProd + " FROM  SanPhamUser INNER JOIN "
		strProd=strProd + " SanPham_User ON SanPhamUser.SanPhamUser_ID = SanPham_User.SanPhamUser_ID INNER JOIN "
		strProd=strProd + " News ON SanPham_User.SanPham_ID = News.NewsID "

		strProd=strProd&" WHERE(DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0)"
		strProd=strProd&" AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0)"

		If (WorkerKSID <> "0") Then 
			strProd = strProd&" AND KiemSoat="&WorkerKSID
		End If
		If (WorkerMHID <> "0") Then 
			strProd = strProd&" AND NhanVienID="&WorkerMHID
		End If
		select case iMaorTenSach 
			case 1
				strTemp	=left(strMaorTenSach,2)
				if isnumeric(strTemp) = true then
					munb=Clng(strTemp) - 1000
				else
					munb = 0
				end if			
				strProd = strProd + " and SanPhamUser_ID = '"& munb &"'"
			case 2
				strProd = strProd + " and News.Title like N'%"& strMaorTenSach &"%'"
			case 3
				strProd = strProd + " and SanPhamUser_Name like N'%"& strMaorTenSach &"%'"
			case 4
				strProd = strProd + " and SanPhamUser_Email like N'%"& strMaorTenSach &"%'"
		end select


		if iOrderBy = 1 then 
			strProd=strProd & " ORDER BY "& strSelSearch &" desc"
		else
			strProd=strProd & " ORDER BY "& strSelSearch 
		end if
		
			

		dim rsProd
		set rsProd=Server.CreateObject("ADODB.Recordset")
		rsProd.open strProd,Con,1
		iSTT=1		 
	%>
	<table width="920" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td colspan="2" align="left" class="CTxtContent">Thời gian:<b> Từ:<%=FromDate%>&nbsp;đến:<%=ToDate%></b></td>
      </tr>
	  	<% 
			if iMaorTenSach = 1 then
			sql11 = "SELECT  Maso, ProviderID,WorkerMuaHangID, AccountingID, DateTime FROM  inputProduct"
			sql11 = sql11 + " where Maso = '"& strMaorTenSach &"'"
			sql11=sql11 + " and AccountingSigna<>0 and StoreSigna<>0 and CreaterSigna<>0 "	
			if NhCC <>0 then
				sql11 = sql11 + " and ProviderID = '"& NhCC &"'"
			end if
			set rsTemp = Server.CreateObject("ADODB.recordset")
			rsTemp.open sql11,Con,1
			do while not rsTemp.eof
				WorkerKSID	=	GetNumeric(rsTemp("AccountingID"),0)
				WorkerMHID	=	GetNumeric(rsTemp("WorkerMuaHangID"),0)
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
      <tr>
        <td align="left" class="CTxtContent">
		<%if NhCC <> 0 then
			Response.Write("Nhà cung cấp:<b>"&getProviderFormID(NhCC)&"</b>")
		end if%>		</td>
        <td align="left" class="CTxtContent">&nbsp;</td>
      </tr>

      <tr>
        <td align="left" class="CTxtContent">
		<%
		if WorkerMHID <> 0 then
			Response.Write("Người mua:<b>"&getNhanVienFromID(WorkerMHID)&"</b>")
		end if%>		</td>
        <td align="left" class="CTxtContent">&nbsp;</td>
      </tr>

      <tr>
        <td width="273" align="left" class="CTxtContent">
		<%	
		if WorkerKSID <> 0 then
			Response.Write("Kiểm soát viên:<b>"&getNhanVienFromID(WorkerKSID)&"<b>") 
		end if
		%>			</td>
        <td align="left" class="CTxtContent">		</td>
      </tr>
    </table>
	<br>
	
<table width="920" border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000" class="CTxtContent">
		<tr align="center" valign="middle" bgcolor="#FFFFFF">
			<td width="2%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>TT</strong></td>
			<% if iMaorTenSach <> 1 then%>
				<td width="3%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>Mã HĐ </strong></td>
			<%
			end if
			
			if NhCC = 0 and iMaorTenSach <> 1 then
			%>	
		    <td width="7%" bgcolor="#FFFFCC" class="CFontVerdana10">Nhà CC </td>
			<%
				end if
			 if iMaorTenSach <> 1 then
			%>
		        <td width="4%" bgcolor="#FFFFCC" class="CFontVerdana10">Ngày</td>
			<%
			end if
			if WorkerKSID = 0  and iMaorTenSach <> 1then
			%>
            <td width="12%" bgcolor="#FFFFCC" class="CFontVerdana10">Kiểm SV </td>
			<%end if
			if WorkerMHID = 0  and iMaorTenSach <> 1 then
			%>
			<td width="13%" bgcolor="#FFFFCC" class="CFontVerdana10">Mua hàng </td>
			<%end if%>
            <td width="40%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>Tên hàng</strong></td>
			<td width="2%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>SL</strong></td>
			<td width="4%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>Đơn giá 
			</strong></td>
			<td width="4%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>VAT
			</strong></td>
			<td width="9%" bgcolor="#FFFFCC" class="CFontVerdana10"><strong>Thành tiền</strong></td>
	  </tr>
	  
		<%
		iTotal =0
		iTotalSL = 0
		iDonGia = 0
		Do while not rsProd.eof
		%>
		<tr  valign="middle" bgcolor="#FFFFFF">
		  <td align="center" width="2%"><%=iSTT%></td>
		  <% if iMaorTenSach <> 1 then%>
		  <td  align="Left" width="3%"><%=rsProd("Maso")%></td>
			<%
			end if
				if NhCC = 0 and iMaorTenSach <> 1 then
			%>	
			<td  align="Left" width="7%"><%=rsProd("ProviderName")%></td>
			<%end if
			if iMaorTenSach <> 1 then
			%>
			<td  align="Left" width="4%"><%=rsProd("DateTime")%></td>
			<%
			end if
			if WorkerKSID = 0  and iMaorTenSach <> 1 then%>
		  <td  align="Left" width="12%"><%=rsProd("Ho_Ten")%></td>
		  <%end if
		  	if WorkerMHID = 0  and iMaorTenSach <> 1 then
		  %>
			<td  align="Left" width="13%"><%=getNhanVienFromID(rsProd("WorkerMuaHangID"))%></td>
			<%end if%>
			<td  align="Left" width="40%"><%=rsProd("Title")%></td>
			<%
			Price 	= 	GetNumeric(rsProd("Price"),0)
			SL		=	GetNumeric(rsProd("Number"),0)	
			iVAT  =  rsProd("VAT")
			TTien = Price*SL
			iDonGia  = iDonGia  + TTien
			iTotalSL = iTotalSL + SL
			TTien = TTien + TTien*iVAT/100	
			iTotal = iTotal + TTien
			%>
			<td  align="center" width="2%"><%=SL%></td>
			<td  align="Right" width="4%"><%=Dis_str_money(Price)%></td>
			<td  align="Center" width="4%"><%=iVAT%></td>

			<td  align="Right" width="9%"><%=Dis_str_money(TTien)%></td>
	    </tr>
		<%
		iSTT=iSTT+1
		rsProd.movenext
		Loop
		%>
</table>
<br>
<br>
    <table width="920" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td width="136" align="right" class="CTxtContent">Tổng số lượng: </td>
        <td width="99" align="left" class="CTxtContent"><%=Dis_str_money(iTotalSL)%> cuốn</td>
        <td width="128" align="right" class="CTxtContent">Giá trước thuế:</td>
        <td width="204" align="left" class="CTxtContent"><b><%=Dis_str_money(iDonGia)&DonviGia%></b></td>
        <td width="258" align="right" class="CTxtContent">Tổng sau thuế:</td>
        <td width="106" align="left" class="CTxtContent"><b><%=Dis_str_money(iTotal)&DonviGia%></b></td>
      </tr>
      <tr>
        <td colspan="6" align="right"class="CTxtContent">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3" align="right" class="CTxtContent">Ghi bằng chữ: </td>
        <td colspan="3" align="left" class="CTxtContent"><b><i><%=tienchu(iTotal)%></i></b></td>
      </tr>
    </table>
<br>

<%End IF%>
    
</body>
</html>
