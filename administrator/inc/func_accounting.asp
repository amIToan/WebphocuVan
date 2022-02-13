<%
function ChungTuPhaSinh(FromDate,ToDate)
FromDate = formatDate()
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
    <td align="center" >&nbsp;</td>
	<td align="center" >&nbsp;</td>
  </tr>	
</table>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
    <td width="2%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
    <td width="3%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>">Số hóa đơn </td>
    <td colspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>">Chứng từ </td>
    <td width="9%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Diễn Giải </strong></td>
    <td colspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">TK</td>
	<td colspan="2"  align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Mã KH </td>
    <td colspan="2"  align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Mã z </td>
    <td width="5%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"> Đơn giá </td>
    <td width="4%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">VAT<br>
    %</td>
    <td width="6%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Thành tiền</td>
    <td colspan="3" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Thông tin người giao dịch </td>
    <td width="6%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Chứng từ kèm </td>
    <td width="4%" rowspan="2" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Tool</td>
  </tr>
  <tr>
    <td width="3%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,1,1)%>"><b>Số</b></td>
    <td width="4%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày  </strong></td>
    <td width="4%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Nợ</td>
    <td width="5%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Có</td>
    <td width="4%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Nợ</strong></td>
    <td width="6%"  align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Có</td>
    <td width="6%"  align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Nợ</td>
    <td width="5%"  align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Có</td>
	<td width="6%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Họ tên </td>
	<td width="8%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Đơn vị </td>
	<%if KhoanPhieuID = 0 and iLoaiPhieu <=1 then%>
    <%end if%>
    <td width="10%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>">Địa chỉ</td>
  </tr>
<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	if KhoanPhieuID > 0 then
		sql	=	"Select * from PhieuKeToan where KhoanPhieuID='"&KhoanPhieuID&"' and "
	else
		sql	=	"Select * from PhieuKeToan where"
	end if
	sql=sql+"  (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu='"&iLoaiPhieu&"') and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 and (iNoiBo= 2 or iNoiBo="&iNoibo&")"
	
	if TKNo <> "" then
		sql=sql + " and TKNo like N'%"& TKNo &"%'"	
	end iF

	if TKCo <> "" then
		sql=sql + " and TKCo like N'%"& TKCo &"%'"	
	end iF
	
	if LapPhieuID <> 0 then
		sql = sql + " and  NVLapPhieuID='"& LapPhieuID &"'"
	end if
	if strLydo	<>"" then 
		sql = sql + " and Lydo like N'%" & strLydo & "%'"
	end if
	if strNguoiNhan	<>"" then 
		sql = sql + " and NguoiGiaoDich like N'%" & strNguoiNhan & "%'"
	end if
	if QuyenThuChi(Session("room")) = 0  then
		sql = sql + " and Username = '"& session("user") &"'"
	end if
	sql=sql+" ORDER BY NgayPhatSinh DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	i = 0
	do while Not rs.eof
		iTien	=	rs("Sotien")+rs("Sotien")*rs("iVAT")/100
		fTien 	=	fTien + iTien
		iSoHD			=	fSoHDCT(rs("ID"),NgayPhatSinh,ifHD)	
%>	
  <tr>
    <td style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(1,1,0,1)%>"><%=iSoHD%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
		<%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=year(rs("NgayPhatSinh"))%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">	<b><%=rs("Name")%></b><br>
	<font  class="CSubTitle">Đ/c: <%=rs("DiaChi")%></font></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=GetNameNV(rs("NVLapPhieuID"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Lydo")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(rs("Sotien"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=rs("iVAT")%></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(iTien)%></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<a href="PhieuThuChi.asp?isEdit=1&ID=<%=rs("ID")%>&endEdit=1"><img src="../../images/bullet277.gif" width="32" height="32" border="0" align="absmiddle"></a>
	<%if iTaiSan = 1 then%>
	<img src="../../images/icons/icon_go_down.gif" border="0" height="15" width="15" onClick="javascript: fChenTS('<%=iSoHD%>','<%=Dis_str_money(rs("Sotien"))%>',<%=rs("iVAT")%>);">	
	<%else%>
<a href="PhieuIn.asp?ID=<%=rs("ID")%>&iLoaiPhieu=<%=iLoaiPhieu%>" target="_blank"><img src="../../images/icons/article.gif" width="15" height="15"  border="0" align="absmiddle"></a> 	
	<%end if%>	</td>
  </tr>

<%
		rs.movenext
	loop
%>	
  <tr>
    <td colspan="5" style="<%=setStyleBorder(1,1,0,1)%>" align="center"> Tổng: </td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%if KhoanPhieuID = 0 then%>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%end if%>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><b><%=Dis_str_money(fTien)&Donvigia%></b></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"></td>
  </tr>
</table>
<%
end function
%>
<%
function TaxOutStore()
	sql	=	"Select * from PhieuKeToan where  "
	sql=sql+"  (DATEDIFF(dd,NgayPhatSinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayPhatSinh,'" & ToDate &"') >= 0) "
	sql=sql+" and (LoaiPhieu=1) and dongy<>0 and ChukyKT<>0 and ChukyTQ<>0 and ChukyLP<>0 and (iNoiBo= 2 or iNoiBo=0) and AttInvoice <> ''"	
	sql=sql+" ORDER BY NgayPhatSinh"
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
%>
<table width="100%"  align="center" CellPadding=0 CellSpacing=0 class="CTxtContent">
	    <TR class="CTieuDeNhoNho" >
	      <TD width="40" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(1,1,1,1)%>">STT</TD>
	      <TD colspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,0)%>">Chúng từ</TD>
	      <TD width="62" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Mã HH</TD>
	      <TD width="569" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Tên hàng hóa</TD>
	      <TD width="66" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Đơn vị</TD>
	      <TD width="89" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">SL</TD>
	      <TD width="89" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Giá bìa </TD>
	      <TD width="54" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Đơn giá </TD>
	      <TD width="131" rowspan="2" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Thành tiền</TD>
  </TR>
	    <TR class="CTieuDeNhoNho" >
	      <TD width="39" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Số</TD>
			<TD width="57" ALIGN="Center" bgcolor="#FFFF99" style="<%=setStyleBorder(0,1,1,1)%>">Ngày</TD>
		</TR>
<%	
	STT = 1
	do while Not rs.eof	
		SoHDKem	= 0
		SoHDKem	=	GetNumeric(replace(Trim(rs("AttInvoice")),"XB",""),0)
		iF SoHDKem > 1000 then
			SanPhamUser_ID	=	SoHDKem	- 1000
			sql="SELECT	* " &_
			"FROM  SanPham_User " &_
			"WHERE    SanPhamUser_ID = '"&SanPhamUser_ID&"' and re_newsid = 0"
			Set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.open sql,con,3
			Do while not rs1.EOF
				SanPham_User_ID = rs1("SanPham_User_ID")
				NewsID		=	rs1("SanPham_ID")
				arSanPhamNhap = GetInfoSanPhamNhap(NewsID)
				idSanPham	=	arSanPhamNhap(0)
				Title		=	arSanPhamNhap(1)
				Soluong		=	FormatNumber(rs1("SanPham_Soluong"),0)
				Giabia		=	FormatNumber(rs1("SanPham_Giabia"),0)
				Gia			=	FormatNumber(rs1("SanPham_Gia"),0)
				tGia		=	CDBl(Gia)*Cdbl(Soluong)
			
			%>
					<TR>
					  <TD ALIGN="Center" style="<%=setStyleBorder(1,1,0,1)%>"><%=STT%></TD>
						<TD ALIGN="Center" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Chungtu")%></TD>
						<TD ALIGN="Center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgayPhatSinh"))%>/<%=month(rs("NgayPhatSinh"))%>/<%=Year(rs("NgayPhatSinh"))%></TD>
						<TD ALIGN="Center" style="<%=setStyleBorder(0,1,0,1)%>"><%=idSanPham%>			</TD>
						<TD ALIGN="Left" style="<%=setStyleBorder(0,1,0,1)%>" >
							<a href="../donhang/print.asp?newsId=<%=NewsID%>"><%= Title %></a>	</TD>
						<TD ALIGN="center" width="66" style="<%=setStyleBorder(0,1,0,1)%>">cuốn</TD>
						<TD ALIGN="Right" width="89" style="<%=setStyleBorder(0,1,0,1)%>"><%= Soluong%></TD>
						<TD ALIGN="Right" width="89" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(Giabia)%></TD>
						<TD ALIGN="Right" width="54" style="<%=setStyleBorder(0,1,0,1)%>"><%= Dis_str_money(Gia) %></TD>
						<TD ALIGN="Right" width="131" style="<%=setStyleBorder(0,1,0,1)%>"><%= Dis_str_money(tGia) %></TD>
					</TR>
					<%
					sTotal = sTotal + tGia
					STT	=	STT+1
					rs1.MoveNext
				Loop
				set rs1 = nothing
			
			end iF
		rs.movenext
	loop
	set rs = nothing
%>
	<TR>
	<TD COLSPAN=9 ALIGN="center" style="<%=setStyleBorder(1,1,0,1)%>">
	<em>Tổng:</em></TD>
	  <TD ALIGN="Right" style="<%=setStyleBorder(0,1,0,1)%>"><em><%= Dis_str_money(sTotal)&DonviGia %></em></TD></TR>
	</TABLE>
<%	

end function
%>