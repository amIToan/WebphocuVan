<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
function  fCTKhoanThu(iCTKhoanThu,fullSigna)
if iCTKhoanThu = 1 then%>
<link href="../../css/styles.css" rel="stylesheet" type="text/css" />
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td colspan="7"  class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(1). Chi tiết khoản thu. </td>
  </tr>
<%end if%>	
<%if iCTKhoanThu = 1 and irbAll <> 0 then%>	
  <tr>
    <td width="4%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>"><b>MS</b></td>
    <td width="6%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày  </strong></td>
    <td width="23%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Đơn vị </strong></td>
    <td width="9%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tổng tiền</strong></td>
    <td width="24%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Lý do </strong></td>
    <td width="16%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Khoản thu</strong> </td>
    <td width="18%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Người thu </strong> </td>
  </tr>
<%end if%>	
<%
	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu = 1) and iNoiBo > 0 "
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if
	if QuyenThuChi(Session("room")) = 0 then
		sql = sql + " and NVThuChiID = '"& GetIDNhanVienUserName(session("user")) &"' "
	end if
	sql=sql+"ORDER BY NVThuChiID DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	iTongThu=0
	do while Not rs.eof
	sotien	=rs("Sotien")
	iVAT	=rs("iVAT")	
	SoTienVAT	=	sotien+sotien*iVAT/100
	iTongThu	=	iTongThu+ GetNumeric(SoTienVAT,0)
	for h = 0 to iCountKTThu
		if arTKThuID(h) = rs("KhoanPhieuID") then
			arValueTKThu(h) = arValueTKThu(h) + GetNumeric(SoTienVAT,0)
		end if
	next
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThuChiID") then
			arNVThu(h) = arNVThu(h) + GetNumeric(SoTienVAT,0)
		end if
	next
%>
<%if iCTKhoanThu = 1 and irbAll <> 0 then%>	
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>">XS00<%=rs("ID")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
		<%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=year(rs("NgayPhatSinh"))%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">	<b><%=rs("Name")%></b><br>
	<font  class="CSubTitle">Đ/c: <%=rs("DiaChi")%></font></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(SotienVAT)%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Lydo")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=GetTKThuChi(rs("KhoanPhieuID"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		if GetNameNV(rs("NVThuChiID")) <> "" then
			Response.Write(GetNameNV(rs("NVThuChiID")))
		else
			Response.Write("Quỹ thu")
		end if
	%></td>
  </tr>

<%end if%>	
<%

		rs.movenext
	loop
%>
<%if iCTKhoanThu = 1 then%>
<%
fTongQuy	=	iTongThu
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp		=	arNVThu(h)
	fTongQuy		=	fTongQuy -	iValueTemp
	if iValueTemp > 0 then
%>
  <tr>
    <td colspan="7" style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> : <%=Dis_str_money(iValueTemp)&Donvigia%></td>
  </tr>
<%
	end if
next
%> 
  <tr>
    <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><font class="CTieuDeNho">Tổng:</font><strong><%=Dis_str_money(iTongThu)&Donvigia%></strong> / Trong đó quỹ thu:<%=Dis_str_money(fTongQuy)%></td>
  </tr>	
</table>
	  
<br>
<%end if
end function
%>

<%
function fCTKhoanChi(iCTKhoanChi,fullSigna)
if iCTKhoanChi = 1 then%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td colspan="7"  class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(2).Chi tiết khoản chi. </td>
  </tr>
  <%end if%>
 <%if iCTKhoanChi = 1 and irbAll <> 0 then%>
  <tr>
    <td width="4%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>"><b>MS</b></td>
    <td width="6%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày  </strong></td>
    <td width="23%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Đơn vị </strong></td>
    <td width="9%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tổng tiền</strong></td>
    <td width="24%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Lý do </strong></td>
    <td width="16%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Khoản chi</strong> </td>
    <td width="18%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Người chi </strong> </td>
  </tr>
<%end if%>	
<%

	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu = 0) and iNoiBo > 0"
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if
	
	if QuyenThuChi(Session("room")) = 0 then
		sql = sql + " and NVThuChiID = '"& GetIDNhanVienUserName(session("user")) &"' "
	end if
	sql=sql+"ORDER BY NVThuChiID DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	iTongChi=0
	do while Not rs.eof
	sotien	=rs("Sotien")
	iVAT	=rs("iVAT")	
	SoTienVAT	=	sotien+sotien*iVAT/100	
	iTongChi	=	iTongChi+ GetNumeric(SoTienVAT,0)
	for h = 0 to iCountKTChi
		if arTKChiID(h) = rs("KhoanPhieuID") then
			arValueTKChi(h) = arValueTKChi(h) + GetNumeric(SoTienVAT,0)
		end if
	next
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThuChiID") then
			arNVChi(h) = arNVChi(h) + GetNumeric(SoTienVAT,0)
		end if
	next

	
%>
<%if iCTKhoanChi = 1 and irbAll <> 0 then%>	
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>">XS00<%=rs("ID")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
		<%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=year(rs("NgayPhatSinh"))%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">	<b><%=rs("Name")%></b><br>
	<font  class="CSubTitle">Đ/c: <%=rs("DiaChi")%></font></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(SoTienVAT)%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Lydo")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=GetTKThuChi(rs("KhoanPhieuID"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">	<%
		if GetNameNV(rs("NVThuChiID")) <> ""  then
			Response.Write(GetNameNV(rs("NVThuChiID")))
		else
			Response.Write("Quỹ chi")
		end if
	%></td>
  </tr>
<%end if%>	
<%

		rs.movenext
	loop
%>
<%if iCTKhoanChi = 1 then%>	
<%
fTongQuy	=	iTongChi
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVChi(h)
	fTongQuy			=	fTongQuy-iValueTemp
	if iValueTemp > 0 then
%>
  <tr>
    <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> : <%=Dis_str_money(iValueTemp)&Donvigia%></td>
  </tr>
<%
	end if
next
%> 
  <tr>
    <td colspan="7" style="<%=setStyleBorder(1,1,0,1)%>"><font class="CTieuDeNho" >Tổng</font>:<b><%=Dis_str_money(iTongChi)&Donvigia%></b>/ Trong đó quỹ chi:<%=Dis_str_money(fTongQuy)%></td>
  </tr>	
</table>
<br>
<%end if
end function
%>


<%
function fCTTamUng(iCTTamUng,fullSigna)
if iCTTamUng = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	<tr>
	  <td colspan="6" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(3).Chi tiết nhận tạm ứng. </td>
  </tr>
 <%end if%>
<%if iCTTamUng = 1  and irbAll <> 0 then%>  
	<tr>
	  <td width="4%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
	  <td width="13%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Mã số </td>
	  <td width="28%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Người nhận </td>
	  <td width="10%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
	  <td width="32%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Lý do </td>
	  <td width="13%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Tiền nhận </td>
    </tr>
<%end if%>	
    <%
	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu= 2) "
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if

	if QuyenThuChi(Session("room")) = 0 then
		sql = sql + " and NVThuChiID = '"& GetIDNhanVienUserName(session("user")) &"' "
	end if
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY NVThuChiID desc"
	else
		sql=sql & " ORDER BY NVThuChiID" 
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1 
	stt =1
	fTTamUng = 0
	Do while not rs.eof
		sotien	=rs("Sotien")
		iVAT	=rs("iVAT")	
		SoTienVAT	=	sotien+sotien*iVAT/100
		for h = 0 to iCountNV
			if arNhanVienID(h) = rs("NVThuChiID") then
				arNVNhanTamUng(h) = arNVNhanTamUng(h) + LamTronTien(SoTienVAT)
			end if
		next
		fTTamUng = fTTamUng + LamTronTien(SoTienVAT)
%>		

		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="../KeToan/PhieuIn.asp?ID=<%=rs("ID")%>&iLoaiPhieu=<%=rs("LoaiPhieu")%>" target="_blank">SX00<%=rs("ID")%></a></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		 <%=getNhanVienFromID(rs("NVThuChiID"))%></td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayPhatSinh"))%>/<%=Month(rs("NgayPhatSinh"))%>/<%=Year(rs("NgayPhatSinh"))%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Lydo")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(LamTronTien(SoTienVAT))%></td>
	    </tr>

<%
		stt = stt+1
		rs.movenext
	loop
	set rs = nothing
%>
<%if iCTTamUng = 1 then%>
<%
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVNhanTamUng(h)
	if iValueTemp > 0 then
%>
  <tr>
 		  <td colspan="6"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> : <b><%=Dis_str_money(iValueTemp)%></b></td>
  </tr>
<%
	end if
next
%> 
		<tr>
		  <td colspan="6"  style="<%=setStyleBorder(1,1,0,1)%>"> <font class="CTieuDeNho"> Tổng</font>:<b><%=Dis_str_money(fTTamUng)%></b></td>
  </tr>
</table>
<br>
<%end if
end function
%>

<%
function fCTNopTamUng(iCTNopTamUng,fullSigna)
if iCTNopTamUng = 1 then%>
   <br> 
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	<tr>
	  <td colspan="6" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(4).Chi tiết nộp tạm ứng. </td>
  </tr>
  <%end if%>
  <%if iCTNopTamUng = 1 and  irbAll<>0 then%>	
	<tr>
	  <td width="4%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
	  <td width="13%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Mã số </td>
	  <td width="28%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Người nhận </td>
	  <td width="10%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
	  <td width="32%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Lý do </td>
	  <td width="13%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Tiền nộp </td>
    </tr>
<%end if%>	
    <%
	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu= 3)"
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if
	
	if QuyenThuChi(Session("room")) = 0 then
		sql = sql + " and NVThuChiID = '"& GetIDNhanVienUserName(session("user")) &"' "
	end if
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY NVThuChiID desc"
	else
		sql=sql & " ORDER BY NVThuChiID" 
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1 
	stt =1
	fTTamUng = 0
	Do while not rs.eof
	sotien	=rs("Sotien")
	iVAT	=rs("iVAT")	
	SoTienVAT	=	sotien+sotien*iVAT/100	
		for h = 0 to iCountNV
			if arNhanVienID(h) = rs("NVThuChiID") then
				arNVNopTamUng(h) = arNVNopTamUng(h) + LamTronTien(SoTienVAT)
			end if
		next
		fTTamUng = fTTamUng + LamTronTien(SoTienVAT)
%>
<%if iCTNopTamUng = 1 and  irbAll<>0 then%>		
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="../KeToan/PhieuIn.asp?ID=<%=rs("ID")%>&iLoaiPhieu=<%=rs("LoaiPhieu")%>" target="_blank">SX00<%=rs("ID")%></a></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		 <%=getNhanVienFromID(rs("NVThuChiID"))%></td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayPhatSinh"))%>/<%=Month(rs("NgayPhatSinh"))%>/<%=Year(rs("NgayPhatSinh"))%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Lydo")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(LamTronTien(SoTienVAT))%></td>
	    </tr>
<%end if%>
<%
		stt = stt+1
		rs.movenext
	loop
	set rs = nothing
%>
<%if iCTNopTamUng = 1 then%>
<%
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVNopTamUng(h)
	if iValueTemp > 0 then
%>
  <tr>
 		  <td colspan="6"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%>:<b><%=Dis_str_money(iValueTemp)%></b></td>
     </tr>
<%
	end if
next
%> 
		<tr>
		  <td colspan="6"  style="<%=setStyleBorder(1,1,0,1)%>"><font  class="CTieuDeNho"> Tổng</font>:<b><%=Dis_str_money(fTTamUng)%></b></td>
     </tr>
</table>
   <br>
<%end if
end function
%>


<%
function fCTThuDH(iCTThuDH)
if iCTThuDH = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td colspan="7" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(5).Chi tiết thu tiền đơn hàng. </td>
  </tr>
<%end if%>
<%if iCTThuDH = 1 and irbAll<>0 then%>  
  <tr> 
    <td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td width="44%" align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tên</b></td>
    <td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Thu tiền </td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Kiểm soát</b> </div></td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tổng tiền </b> </div></td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày TT </td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Ngày đặt</b> </td>
  </tr>
<%end if%>     
<%
 'Nhân viên thu tiền sách
Set rs=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM SanPhamUser " 
sql=sql + " where SanPhamUser_Status= 2 "
if QuyenThuChi(Session("room")) = 0 then
	sql = sql + " and NVThutienID = '"& GetIDNhanVienUserName(session("user")) &"' "
end if
sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') <= 0) "
sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & ToDate & "') >= 0)"
sql=sql + " ORDER BY NVThutienID" 
rs.open sql,con,3
stt = 1
iCuocNetco =0
iCuocBuuDien=0
fTongTienThuDH = 0
Do while not rs.eof 
	SanPhamUser_ID		=	rs("SanPhamUser_ID")
	SanPhamUser_Name	=	rs("SanPhamUser_Name")
	SanPhamUser_Email	=	rs("SanPhamUser_Email")
	SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
	SanPhamUser_Address	=	rs("SanPhamUser_Address")
	NgayXuLy			=	rs("NgayXuLy")
	NgayTT				=	rs("NgayThanhToan")
	strCMND				=	rs("CMND")
	KSoat				=	getNhanVienFromID(rs("KiemSoat"))
	NVGiaoHang			=	getNhanVienFromID(rs("NhanVienID"))	
	NVThutien			=	getNhanVienFromID(rs("NVThutienID"))		
	iTien 	= 	LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND))
	fTongTienThuDH		=	fTongTienThuDH+iTien
	if 	NVThutien = "Netco" then
		iTien	=	iTien	-	GetCuocBuuDienThucID(SanPhamUser_ID)
		iCuocNetco = iCuocNetco + GetCuocBuuDienThucID(SanPhamUser_ID)
	end if
	if NVThutien = "Bưu điện" then
		iTien	=	iTien	-	GetCuocBuuDienThucID(SanPhamUser_ID)
		iCuocBuuDien = iCuocBuuDien + GetCuocBuuDienThucID(SanPhamUser_ID)
	end if
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThutienID") then
			arNVThuDH(h) = arNVThuDH(h) + iTien
		end if
	next	
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThutienID") and NgayTT <> #1/1/1900# and isdate(NgayTT) = true then
			arNVTToanDH(h) = arNVTToanDH(h) + iTien
		end if
	next	
%>
<%if iCTThuDH = 1 and irbAll<>0 then%>
  <tr <%if stt mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td width="4%"align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		<a href="../thongke/ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
    <%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a></td>
    <td align="left" valign="middle" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%><br>
	<font class="CSubTitle">
	<i>Địa chỉ</i>: <%=SanPhamUser_Address%></font>	</td>
    <td width="14%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=NVThutien%></td>
	<td width="13%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%></td>
	<td width="10%"  style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTien)%></td>
	<td width="7%" style="<%=setStyleBorder(0,1,0,1)%>" align="center">
	<%if NgayTT <> #1/1/1990# or isdate(NgayTT)= true then%>
		<%=Day(ConvertTime(NgayTT))%>/<%=Month(ConvertTime(NgayTT))%>/<%=Year(ConvertTime(NgayTT))%>	
	<%else%>
		<img src="../images/icon-banner-new.gif" height="16" width="16" border="0">
	<%end if%>	</td>
	<td width="8%" style="<%=setStyleBorder(0,1,0,1)%>" align="center">
		<%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%>	</td>
  </tr>


<%
end if
	stt=stt + 1
	rs.movenext
loop
%>
<%if iCTThuDH = 1 then%>
<%
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVThuDH(h)
	if iValueTemp > 0 then
%>
  <tr>
 		  <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> :<b><%=Dis_str_money(iValueTemp)%></b></td>
  </tr>
<%
	end if
next
%> 
  <tr >
    <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><font class="CTieuDeNho"> Tổng</font>:<b><%=Dis_str_money(fTongTienThuDH)&Dovigia%></b></td>
  </tr>
</table>
<br>
<%end if
end function
%>


<%
function fCTChiDH(iCTChiDH,fullSigna)
if iCTChiDH = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	<tr>
	  <td colspan="7" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(6).Chi tiết nhập hàng. </td>
  </tr>
  <%end if%>
  <%if iCTChiDH = 1  and irbAll <> 0 then%>	
	<tr>
	  <td width="6%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
	  <td width="30%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Số</td>
	  <td width="7%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
	  <td width="16%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Người mua </td>
	  <td width="14%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Thanh Toán </td>
	  <td width="2%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"> SL </td>
	  <td width="11%"align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">T.Tiền</td>
	</tr>
<%end if%>	
 <%
 	' Nhân viên thu tiền sách
	sql ="SELECT  inProductID,Maso,ProviderName,Ho_Ten,WorkerThanhToanID,AccountingID,DateTime FROM  inputProduct INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID INNER JOIN Nhanvien ON inputProduct.WorkerMuaHangID = Nhanvien.NhanVienID "
	sql = sql + "where (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0) "
	if QuyenThuChi(Session("room")) = 0 then
		sql = sql + " and WorkerThanhToanID = '"& GetIDNhanVienUserName(session("user")) &"' "
	end if	
	if fullSigna = false then
		sql=sql+" and (inputProduct.AccountingSigna<>0 or inputProduct.StoreSigna<>0 or inputProduct.CreaterSigna<>0) "
	else
		sql=sql+" and inputProduct.AccountingSigna<>0 and inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0 "
	end if	
	
	if iOrderBy = 1 then 
		sql=sql & " ORDER BY DateTime desc"
	else
		sql=sql & " ORDER BY DateTime" 
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1
	stt	=	1 
	fTongChiDH = 0
	Do while not rs.eof
		for h = 0 to iCountNV
			if arNhanVienID(h) = rs("WorkerThanhToanID") then
				arNVChiDH(h) = arNVChiDH(h) + LamTronTien(GetTTien(rs("inProductID")))
			end if
		next
		fTongChiDH	=	fTongChiDH+LamTronTien(GetTTien(rs("inProductID")))
if iCTChiDH = 1  and irbAll <> 0 then%>		
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
		  <a href="../thongke/Report_SoHD.asp?inProductID=<%=rs("inProductID")%>" target="_parent" class="CSubMenu"><%=rs("Maso")%></a><br>
		  <font class="CSubTitle">
		  <i>Nhà cung cấp</i>: <%=rs("ProviderName")%>		  </font>		  </td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("DateTime"))%>/<%=Month(rs("DateTime"))%>/<%=Year(rs("DateTime"))%></td>
		  <td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ho_Ten")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%=getNhanVienFromID(rs("WorkerThanhToanID"))%>		  </td>
		  <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetTotalSPinHD(rs("inProductID"))%></td>
		  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(LamTronTien(GetTTien(rs("inProductID"))))%></td>
		</tr>

<%end if	
		stt	=	stt+1
		rs.movenext
	loop
%>
<%if iCTChiDH = 1 then%>
<%
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVChiDH(h)
	if iValueTemp > 0 then
%>
  <tr>
 		  <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> : <b><%=Dis_str_money(iValueTemp)%></b></td>
  </tr>
<%
	end if
next
%> 
		<tr>
		  <td colspan="7"  style="<%=setStyleBorder(1,1,0,1)%>"><font class="CTieuDeNho"> Tổng</font>:<b><%=Dis_str_money(fTongChiDH)&Donvigia%></b></td>
  </tr>
</table>
<br>
<%end if
end function
%>	


<%
function fCTNoCu(iCTNoCu)
if iCTNoCu = 1 then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td colspan="7" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(8).Chi tiết thanh toán nợ cũ. </td>
  </tr>
<%end if%>  
<%if iCTNoCu = 1 and irbAll <>0 then%>  
  <tr> 
    <td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td width="44%" align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tên</b></td>
    <td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Thu tiền </td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Kiểm soát</b> </div></td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Tổng tiền </b> </div></td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>">Ngày TT </td>
	<td align="center" bgcolor="#FFFFCC" class="CTextStrong" style="<%=setStyleBorder(0,1,1,1)%>"><b>Ngày đặt</b> </td>
  </tr>
<%end if%>     
<%
 'Nhân viên thu tiền sách
Set rs=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM SanPhamUser " 
sql=sql + " where SanPhamUser_Status= 2 "
sql=sql + " AND (DATEDIFF(dd, NgayXuLy, '" & FromDate & "') > 0) "
sql=sql + " AND (DATEDIFF(dd, NgayThanhToan, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, NgayThanhToan, '" & ToDate & "') >= 0) "
if QuyenThuChi(Session("room")) = 0 then
	sql = sql + " and NVThutienID = '"& GetIDNhanVienUserName(session("user")) &"' "
end if
sql=sql + " ORDER BY NVThutienID" 
rs.open sql,con,3
stt = 1
iCuocNetco =0
iCuocBuuDien=0
fTongNocu=0
Do while not rs.eof 
	SanPhamUser_ID		=	rs("SanPhamUser_ID")
	SanPhamUser_Name	=	rs("SanPhamUser_Name")
	SanPhamUser_Email	=	rs("SanPhamUser_Email")
	SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
	SanPhamUser_Address	=	rs("SanPhamUser_Address")
	NgayXuLy			=	rs("NgayXuLy")
	NgayTT				=	rs("NgayThanhToan")
	strCMND				=	rs("CMND")
	KSoat				=	getNhanVienFromID(rs("KiemSoat"))
	NVGiaoHang			=	getNhanVienFromID(rs("NhanVienID"))	
	NVThutien			=	getNhanVienFromID(rs("NVThutienID"))		
	iTien 	= 	LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND))
	if 	NVThutien = "Netco" then
		iTien	=	iTien	-	GetCuocBuuDienThucID(SanPhamUser_ID)
		iCuocNetco = iCuocNetco + GetCuocBuuDienThucID(SanPhamUser_ID)
	end if
	if NVThutien = "Bưu điện" then
		iTien	=	iTien
		iCuocBuuDien = iCuocBuuDien + GetCuocBuuDienThucID(SanPhamUser_ID)
	end if
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NVThutienID") then
			arNVTToanNoCu(h) = arNVTToanNoCu(h) + iTien
		end if
	next
	fTongNocu	=	fTongNocu+	iTien
	
%>
<%if iCTNoCu = 1 and irbAll <>0 then%>
  <tr <%if stt mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td width="4%"align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		<a href="../thongke/ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
    <%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a></td>
    <td align="left" valign="middle" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%><br>
	<font class="CSubTitle">
	<i>Địa chỉ</i>: <%=SanPhamUser_Address%></font>	</td>
    <td width="14%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=NVThutien%></td>
	<td width="13%"  style="<%=setStyleBorder(0,1,0,1)%>"><%=KSoat%></td>
	<td width="10%"  style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTien)%></td>
	<td width="7%" style="<%=setStyleBorder(0,1,0,1)%>" align="center">
	<%if NgayTT <> #1/1/1990# or isdate(NgayTT)= true then%>
		<%=Day(ConvertTime(NgayTT))%>/<%=Month(ConvertTime(NgayTT))%>/<%=Year(ConvertTime(NgayTT))%>	
	<%else%>
		<img src="../images/icon-banner-new.gif" height="16" width="16" border="0">
	<%end if%>	</td>
	<td width="8%" style="<%=setStyleBorder(0,1,0,1)%>" align="center">
		<%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%>	</td>
  </tr>
<%
end if
	stt=stt + 1
	rs.movenext
loop
%>
<%if iCTNoCu = 1 then%>
  <tr>
    <td colspan="7" align="left" style="<%=setStyleBorder(1,1,0,1)%>"><%
		for h = 0 to iCountNV -1
		 	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
			TTNoDH			=	arNVTToanNoCu(h)
			if TTNoDH > 0 then
				Response.Write(strNhanVien&": "&Dis_str_money(TTNoDH)&" đ <br>")
			end if
		next
	%></td>
  </tr>
  <tr>
  	<td colspan="7" style="<%=setStyleBorder(1,1,0,1)%>">
	<font class="CTieuDeNho"> Tổng:</font>
		<%=Dis_str_money(fTongNocu)%> đ
	</td>
  </tr>
</table>
<br>
<%end if
end function
%>


<%
function fCTThuDHVAT(iCTThuDHVAT,fullSigna)
if iCTThuDHVAT = 1 and QuyenThuChi(Session("room")) <> 0 then%>	  
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td colspan="11"  class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,0)%>">Chi tiết khoản thu có VAT. </td>
    </tr>
  <tr>
    <td width="3%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>"><b>STT</b></td>
    <td colspan="3" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Hóa đơn chứng từ biên lai nộp thuế </strong><strong></strong></td>
    <td width="23%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên người mua </strong></td>
    <td width="7%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã số thuế</strong></td>
    <td width="19%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mặt hàng </strong> </td>
    <td width="8%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Doanh số mua chưa có thuế </strong></td>
    <td width="3%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thuế suất (%) </strong></td>
    <td width="7%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thuế GTGT </strong></td>
    <td width="8%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ghi chú </strong></td>
  </tr>
  <tr>
    <td width="8%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Ký hiệu HĐ </td>
    <td width="7%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Số hóa đơn </td>
    <td width="7%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
    </tr>
<%
	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu = 1)  and (iNoiBo = 0  or iNoiBo = 2)"
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if	
	sql=sql+"ORDER BY NVThuChiID DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	iTongThuVAT=0
	stt=0
	iTongThuSauVAT = 0
	do while Not rs.eof
	stt=	stt+1
	iTongThuVAT	=	iTongThuVAT+ GetNumeric(rs("Sotien"),0)
%>
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>" align="center"><%=stt%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Kemtheo")%>		</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Chungtu")%></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=year(rs("NgayPhatSinh"))%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Name")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("MST")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<%=rs("Lydo")%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(rs("Sotien"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("iVAT")%>%</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
	<%
		iTien 	=	GetNumeric(rs("Sotien"),0)
		VAT		=	GetNumeric(rs("iVAT"),0)
		iValueTemp	=	iTien*VAT/100
		iValueTemp	=	Round(iValueTemp)
		iTongThuSauVAT = iTongThuSauVAT + iValueTemp
		Response.Write(Dis_str_money(iValueTemp))
	%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ghichu")%></td>
  </tr>
<%

		rs.movenext
	loop
%>

  <tr>
    <td colspan="7" align="center" class="CTieuDeNho" style="<%=setStyleBorder(1,1,0,1)%>">Tổng:</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTongThuVAT)&Donvigia%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTongThuSauVAT)&Donvigia%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
  </tr>	
</table>
<br>
<%end if
end function
%>

<%
function fCTChiDHVAT(iCTChiDHVAT,fullSigna)
if iCTChiDHVAT = 1 and QuyenThuChi(Session("room")) <> 0 then%>	  
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
  <tr>
    <td colspan="11"  class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,0)%>">Chi tiết khoản chi có VAT. </td>
    </tr>
  <tr>
    <td width="3%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>"><b>STT</b></td>
    <td colspan="3" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Hóa đơn chứng từ biên lai nộp thuế </strong><strong></strong></td>
    <td width="23%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên người mua </strong></td>
    <td width="7%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mã số thuế</strong></td>
    <td width="19%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Mặt hàng </strong> </td>
    <td width="8%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Doanh số mua chưa có thuế </strong></td>
    <td width="3%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thuế suất (%) </strong></td>
    <td width="7%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thuế GTGT </strong></td>
    <td width="8%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ghi chú </strong></td>
  </tr>
  <tr>
    <td width="8%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Ký hiệu HĐ </td>
    <td width="7%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Số hóa đơn </td>
    <td width="7%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</td>
    </tr>
<%
	sql	=	"Select * from PhieuKeToan"
	sql=sql+" where (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu = 0)   and (iNoiBo = 0  or iNoiBo = 2)"
	if fullSigna = false then
		sql=sql+" and (dongy<>0 or ChukyKT<>0 or ChukyTQ<>0 or ChukyLP<>0) "
	else
		sql=sql+" and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 "
	end if	
	sql=sql+"ORDER BY NVThuChiID DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	iTongChiVAT=0
	stt=0
	iTongChiSauVAT = 0
	do while Not rs.eof
	stt=	stt+1
	iTongChiVAT	=	iTongChiVAT+ GetNumeric(rs("Sotien"),0)
%>
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>" align="center"><%=stt%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Kemtheo")%>		</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Chungtu")%></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=year(rs("NgayPhatSinh"))%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Name")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;<%=rs("MST")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<%=rs("Lydo")%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(rs("Sotien"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("iVAT")%>%</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
	<%
		iTien 	=	GetNumeric(rs("Sotien"),0)
		VAT		=	GetNumeric(rs("iVAT"),0)
		iValueTemp	=	iTien*VAT/100
		iValueTemp	=	Round(iValueTemp)
		iTongChiSauVAT = iTongChiSauVAT + iValueTemp
		Response.Write(Dis_str_money(iValueTemp))
	%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;<%=rs("Ghichu")%></td>
  </tr>
<%

		rs.movenext
	loop
%>

  <tr>
    <td colspan="7" align="center" class="CTieuDeNho" style="<%=setStyleBorder(1,1,0,1)%>">Tổng:</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTongChiVAT)&Donvigia%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTongChiSauVAT)&Donvigia%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
  </tr>	
</table>
<br>
<%end if
end function
%>

<%
function fCTTraNCC(iCTTraNCC,fullSigna)
%>
<%if iCTTraNCC = 1 then%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td colspan="8" class="CTieuDeNho" style="<%=setStyleBorder(1,1,1,1)%>">(7) Chi tiết trả nhà cung cấp.</td>
  </tr>
  <%end if%>
  <%if iCTTraNCC = 1  and irbAll <> 0 then%>	  
  <tr>
    <td width="8%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(1,1,1,1)%>" align="center"><strong>Mã số</strong></td>
    <td width="15%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Nhà cung cấp</strong></td>
    <td width="12%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong><strong>Người trả</strong></strong></td>
    <td width="7%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Ngày trả</strong></td>
    <td width="39%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Tên sách</strong></td>
    <td width="7%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Giá nhập</strong><br>
    <font class="CSubTitle">(Gồm VAT)</font></td>
    <td width="4%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Số lượng </strong></td>
    <td width="8%" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>" align="center"><strong>Thành tiền </strong></td>
  </tr>
  <%end if%>
<%
sql = "SELECT  Nhanvien.Ho_Ten, TraSach.IDTraNCC, TraSach.ProductID, TraSach.SLTraNCC, TraSach.NhanVienID,TraSach.NgayTra,"
sql =	sql + " Product.NewsID, Product.Number, Product.Giabia, Product.Price, Product.VAT, inputProduct.Maso, Provider.ProviderName, "
sql =	sql + " SanPhamNhap.Title, SanPhamNhap.Tacgia, SanPhamNhap.nxb"
sql =	sql + " FROM Product INNER JOIN TraSach ON Product.ProductID = TraSach.ProductID INNER JOIN "
sql =	sql + " inputProduct ON Product.inProductID = inputProduct.inProductID INNER JOIN "
sql =	sql + " Provider ON inputProduct.ProviderID = Provider.ProviderID INNER JOIN "
sql =	sql + " Nhanvien ON TraSach.NhanVienID = Nhanvien.NhanVienID INNER JOIN "
sql =	sql + " SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID"
sql =	sql + " where (DATEDIFF(dd,NgayTra,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayTra,'" & ToDate &"') >= 0) "

if fullSigna = false then
	sql	=	sql+" and ( inputProduct.AccountingSigna<>0 or inputProduct.StoreSigna<>0 or inputProduct.CreaterSigna<>0) "
else
	sql	=	sql+"  and inputProduct.AccountingSigna<>0 and inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0 "
end if
sql =	sql + " ORDER BY Nhanvien.Ho_Ten"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,con,1
iTongTraNCC = 0
iTongSL	=	0
do while not rs.eof
	iTien 	=	GetNumeric(rs("Price"),0)
	VAT		=	GetNumeric(rs("VAT"),0)
	SLTraNCC=	GetNumeric(rs("SLTraNCC"),0)
	iTongSL	=	iTongSL+SLTraNCC
	iValueTemp	=	iTien	+ iTien*VAT/100
	iValueTemp	=	iValueTemp*SLTraNCC
	iTongTraNCC	=	iTongTraNCC+iValueTemp
	for h = 0 to iCountNV
		if arNhanVienID(h) = rs("NhanVienID") then
			arNVTraNCC(h) = arNVTraNCC(h) + iValueTemp
		end if
	next
%>
  <%if iCTTraNCC = 1  and irbAll <> 0 then%>  
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>"><%=rs("Maso")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("ProviderName")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ho_Ten")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayTra"))%>/<%=month(rs("NgayTra"))%>/<%=Year(rs("NgayTra"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Title")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iTien)%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=SLTraNCC%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(iValueTemp)%></td>
  </tr>
  <%end if%> 
<%
	rs.movenext
loop
%>  

<%if iCTTraNCC = 1 then%> 
  
<%
for h = 0 to iCountNV -1
	strNhanVien  	=  getNhanVienFromID(arNhanVienID(h))
	iValueTemp			=	arNVTraNCC(h)
	if iValueTemp > 0 then
%>
  <tr>
 		  <td colspan="8"  style="<%=setStyleBorder(1,1,0,1)%>"><%=strNhanVien%> :<b><%=Dis_str_money(iValueTemp)%></b></td>
  </tr>

<%
	end if
next
%> <tr>
    <td colspan="8"  style="<%=setStyleBorder(1,1,0,1)%>"><font class="CTieuDeNho">Tổng:</font>Tổng tiền:<%=Dis_str_money(iTongTraNCC)&Donvigia%>/<%=iTongSL%> cuốn</td>
  </tr>
</table>
<br>
<%
end if
end function
%>

