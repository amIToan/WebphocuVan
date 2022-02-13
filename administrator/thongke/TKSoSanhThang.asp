<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_Datetime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Thang2	=	Month(now())
	Nam2	=	Year(now())
	Thang1	=	0
	Nam1	=	0
	iRaOderBy	=	GetNumeric(Request.Form("RaOderBy"),0)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/styles.css" rel="stylesheet" type="text/css">

<title>THỐNG KÊ SO SÁNH THEO THÁNG</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
if Clng(Request.Form("reThongKe")) <> 1 then
	Title_This_Page="Thống kê -> So sách theo tháng"
	Call header()
	Call Menu()
	

	
%>
<form id="frmReportMH" name="frmReportMH" method="post" action="TKSoSanhThang.asp">
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td  background="../../images/T1.jpg" height="20"></td>
  </tr>
  <tr>
    <td background="../../images/t2.jpg"><table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
      
      <tr>
        <td width="100%" height="29" colspan="3" align="center" bordercolor="#FFFFFF"class="CTxtContent">Thống kê theo: &nbsp;Giản đơn
          <input name="RaOderBy" type="radio" value="0" checked <%if iOrderBy =0 then Response.Write("checked") end if%>>
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif">/ Đầy đủ
                <input name="RaOderBy" type="radio" value="1" <%if iOrderBy =1 then Response.Write("checked") end if%>>
            </font></font></span></td>
      </tr>
      <tr>
        <td height="29" colspan="3" align="center" bordercolor="#FFFFFF" class="CTxtContent">Từ tháng:
          <%
					Call List_Month_WithName(Thang1,"MM","Thang1")
					Call List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
				%>
              <img src="../images/right.jpg" width="9" height="9" align="absmiddle" />
			  Đến tháng
              <%
					Call List_Month_WithName(Thang2,"MM","Thang2")
					Call List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
				%>
				<input name="reThongKe" type="hidden" value="1">
				<input type="submit" name="Submit" value="Thông kê">
			  </font>
              </p></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td background="../../images/T3.jpg" height="8"></td>
  </tr>
</table>
</form>
<%Call Footer()%>
<%else%>
<%
	Ngay1			=	1
	Thang1			=	GetNumeric(Request.form("Thang1"),0)
	Nam1			=	GetNumeric(Request.form("Nam1"),0)
	Ngay2			=	day(now())
	Thang2			=	GetNumeric(Request.form("Thang2"),0)
	Nam2			=	GetNumeric(Request.form("Nam2"),0)
	
	arTemp	=	KhoangThang(Thang1,Nam1,Thang2,Nam2)
	iSoThang = Ubound(arTemp,2)-1	
	Redim arSoDH(iSoThang+1)
	Redim arTongThu(iSoThang+1)
	Redim arTongLai(iSoThang+1)
	for k = 0 to iSoThang
		iNgay1			=	1
		TempThang			=	arTemp(0,k)
		TempNam			=	arTemp(1,k)
		iNgay2			=	getDayInMonth(TempThang,TempNam)

		FromDate=TempThang & "/" & iNgay1 & "/" & TempNam
		ToDate=TempThang & "/" & iNgay2 & "/" & TempNam
	
		FromDate=FormatDatetime(FromDate)
		ToDate=FormatDatetime(ToDate)
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="SELECT SanPhamUser_ID,SanPhamUser_Status,CMND FROM SanPhamUser " 
		if iRaOderBy = 1 then
		sql=sql + " where SanPhamUser_Status <> 0 "
		else
		sql=sql + " where (SanPhamUser_Status = 2 or SanPhamUser_Status=14 or SanPhamUser_Status=1 or SanPhamUser_Status=6 or SanPhamUser_Status = 15) "
		end if
		sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
		sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
		rs.open sql,con,3
		STT = 0
		fTongTienXuat =0
		fTongTienThu =0
		Do while not rs.eof 
			SanPhamUser_ID		=	rs("SanPhamUser_ID")
			strCMND				=	rs("CMND")
			sql =       "SELECT XuatKho.SoLuong,Product.Price,Product.VAT "
			sql = sql + " FROM SanPham_User INNER JOIN SanPhamUser ON SanPham_User.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID"
			sql = sql + " INNER JOIN XuatKho ON SanPham_User.SanPham_User_ID = XuatKho.SanPham_User_ID"
			sql = sql + " INNER JOIN Product ON XuatKho.ProductID = Product.ProductID"
			sql = sql + " WHERE SanPhamUser.SanPhamUser_ID = '"&SanPhamUser_ID&"'  and SanPham_User.re_newsid = 0"
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
			stt=stt + 1		
			rs.movenext
		loop	
		arSoDH(k)=stt
		arTongThu(k)=	fTongTienThu
		arTongLai(k)=	fTongTienThu - fTongTienXuat
	next
	set rs = nothing
		
%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
    <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
        <em>ĐT: <%=soDT%>  - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%>   </em></div></td>
  </tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" class="author"><div align="center">
      <p>&nbsp;</p>
      <p>THỐNG KÊ DỮ LIỆU THEO THÁNG </p>
    </div></td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">
	
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" align="center" valign="bottom"><table border="0" cellspacing="0" cellpadding="0" class="CTxtContent">
  <tr>
    <td valign="top">Số<br>ĐH</td>
  	<td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
<%for i = 0 to iSoThang%>
    <td width="25" align="center" valign="bottom">
	<%=arSoDH(i)%><br>
	<%for k = 0 to arSoDH(i) step 2
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
	next%>	</td>
<%next%>
	<td width="25" align="center" valign="bottom">	</td>
  </tr>
  <tr>
    <td align="right" valign="top">&nbsp;</td>
  	<td valign="top"><br></td>
 	<%for i = 0 to iSoThang%>
  	<td style="border-top:#000000 solid 1" align="center" valign="top">
		<%=arTemp(0,i)%>	</td>
	<%next%>
	<td style="border-top:#000000 solid 1" align="right" valign="top" width="50">
	Tháng  	</td>
  </tr>
</table></td>
        <td width="50%" align="center" valign="bottom">
		<%
	Redim arSoDHHuy(iSoThang+1)
	for k = 0 to iSoThang
		iNgay1			=	1
		TempThang			=	arTemp(0,k)
		TempNam			=	arTemp(1,k)
		iNgay2			=	getDayInMonth(TempThang,TempNam)

		FromDate=TempThang & "/" & iNgay1 & "/" & TempNam
		ToDate=TempThang & "/" & iNgay2 & "/" & TempNam
		
		FromDate=FormatDatetime(FromDate)
		ToDate=FormatDatetime(ToDate)
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="SELECT SanPhamUser_ID,SanPhamUser_Status,CMND FROM SanPhamUser " 
		sql=sql + " where SanPhamUser_Status= 3 "
		sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
		sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
		rs.open sql,con,3
		if not rs.eof then
			arSoDHHuy(k)=rs.recordcount - 1
		else
			arSoDHHuy(k)= 0
		end if
		
		set rs=nothing
	next		
		%>
		
		<table border="0" cellspacing="0" cellpadding="0" class="CTxtContent">
  <tr>
    <td valign="top">Số<br>ĐH</td>
  	<td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
<%for i = 0 to iSoThang%>
    <td width="25" align="center" valign="bottom">
	<%=arSoDHHuy(i)%><br>
	<%for k = 0 to arSoDHHuy(i) step 2
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
	next%>	</td>
<%next%>
	<td width="25" align="center" valign="bottom">	</td>
  </tr>
  <tr>
    <td align="right" valign="top">&nbsp;</td>
  	<td valign="top"><br></td>
 	<%for i = 0 to iSoThang%>
  	<td style="border-top:#000000 solid 1" align="center" valign="top">
		<%=arTemp(0,i)%>	</td>
	<%next%>
	<td style="border-top:#000000 solid 1" align="right" valign="top" width="50">
	Tháng  	</td>
  </tr>
</table>		</td>
      </tr>
      <tr>
        <td align="center" class="CFontVerdana10">Đơn hàng thành công</td>
        <td class="CFontVerdana10" align="center">Đơn hàng hủy bỏ </td>
      </tr>
    </table>
	

	</td>
  </tr>
  <tr>
    <td align="center" class="CTxtVerdana10Weight"><p class="CFontVerdana10">&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p></td>
  </tr>
  <tr>
    <td align="center">
	<table border="0" cellspacing="0" cellpadding="0" class="CTxtContent">
  <tr>
    <td valign="top" align="right">
	Triệu	</td>
  	<td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
<%for i = 0 to iSoThang%>
    <td width="50" align="center" valign="bottom">
	<%=round(arTongThu(i)/1000000,3)%><br>
	<%for k = 0 to round(arTongThu(i)/1000000)
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
	next%>	</td>
<%next%>
	<td width="25" align="center" valign="bottom">	</td>
  </tr>
  <tr>
    <td align="right" valign="top">&nbsp;</td>
  	<td valign="top"><br></td>
 	<%for i = 0 to iSoThang%>
  	<td style="border-top:#000000 solid 1" align="center" valign="top">
		<%=arTemp(0,i)%>	</td>
	<%next%>
	<td style="border-top:#000000 solid 1" align="right" valign="top" width="50">
  		Tháng  	</td>
  </tr>
</table>	</td>
  </tr>
  <tr>
    <td align="center" class="CTxtVerdana10Weight"><p class="CFontVerdana10">Thống kê tổng doanh số theo tháng </p>
    <p>&nbsp;</p>
    <p>&nbsp;</p></td>
  </tr>
  <tr>
    <td align="center">

	<table border="0" cellspacing="0" cellpadding="0" class="CTxtContent">
  <tr>
    <td valign="top" align="right">
	Triệu	</td>
  	<td valign="top" style="border-right:#000000 solid 1">&nbsp;</td>
<%for i = 0 to iSoThang%>
    <td width="50" align="center" valign="bottom">
	<%=round(arTongLai(i)/1000000,3)%><br>
	<%for k = 0 to round(arTongLai(i)/1000000)
		Response.Write("<img src=""../../images/TKe.jpg"" width=""20"" height=""1""><br>")
	next%>	</td>
<%next%>
	<td width="25" align="center" valign="bottom">	</td>
  </tr>
  <tr>
    <td align="right" valign="top">&nbsp;</td>
  	<td valign="top"><br></td>
 	<%for i = 0 to iSoThang%>
  	<td style="border-top:#000000 solid 1" align="center" valign="top">
		<%=arTemp(0,i)%>	</td>
	<%next%>
	<td style="border-top:#000000 solid 1" align="right" valign="top" width="50">
  		Tháng  	</td>
  </tr>
</table>	</td>
  </tr>
  <tr>
    <td align="center" class="CFontVerdana10">Thông kê lợi nhuận theo tháng </td>
  </tr>
</table>




<%end if%>

</body>
</html>
<%
function getDayInMonth(iMonth,iYear)
	iSoNgay = 0
	select case iMonth
		case 1
			iSoNgay = 31
		case 2
			if iYear mod 4 = 0 then
				iSoNgay = 29
			else
				iSoNgay = 28
			end if
		case 3
			iSoNgay = 31
		case 4
			iSoNgay = 30
		case 5
			iSoNgay = 31
		case 6
			iSoNgay = 30
		case 7
			iSoNgay = 31
		case 8	
			iSoNgay = 31
		case 9
			iSoNgay = 30
		case 10
			iSoNgay = 31
		case 11
			iSoNgay = 30
		case 12
			iSoNgay = 31
	end select
	getDayInMonth = iSoNgay
end function

function KhoangThang(Thang1,Nam1,Thang2,Nam2)

	if Nam2 < Nam1 then
		Nam2 = Nam1
	end if
	if Nam2 = Nam1 and Thang1 > Thang2 then
		Thang1 = Thang2
	end if
	k = 0
	TempThang1	=	Thang1
	TempNam1	=	Nam1
	TempThang2	=	Thang2
	TempNam2	= 	Nam2
	Do while (Thang1 <> Thang2) or (Nam1 <> Nam2) 
		if Thang1 > 12 then
			Nam1 = Nam1 + 1
			Thang1 = 1
		end if
		k = k+1
		Thang1 = Thang1 + 1				
	loop	
	redim 	arThangNam(1,k+1)
	k	= 0
	Do while (TempThang1 <> TempThang2) or (TempNam1 <> TempNam2) 
		if TempThang1 > 12 then
			TempNam1 = TempNam1 + 1
			TempThang1 = 1
		end if
		
		arThangNam(0,k) = TempThang1
		arThangNam(1,k) = TempNam1
		k=k+1
		TempThang1 = TempThang1 + 1
	loop
	arThangNam(0,k) = TempThang1
	arThangNam(1,k) = TempNam1
	KhoangThang = arThangNam
end function
%>