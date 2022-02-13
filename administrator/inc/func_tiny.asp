<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<%
sub selLoaiSP(stt)
%>
<link href="../../css/styles.css" rel="stylesheet" type="text/css" />
<select name="selLSP<%=stt%>" size="1">
	<option value="0" selected="selected">Lựa chọn</option>
	<%
	strSel="Select CategoryID,CategoryName,YoungestChildren From NewsCategory where CategoryLoai=3 and categoryLevel = 1"
	dim rsNewCat
	set rsNewCat=Server.CreateObject("ADODB.Recordset")
	rsNewCat.open strSel,Con,1 
	%>
	<%
	Dim arChilds 
	Do while not rsNewCat.eof
	%>
		<option value="<%=rsNewCat("CategoryID")%>">*<%=rsNewCat("CategoryName")%></option>
	<%
		if rsNewCat("YoungestChildren") >0 then
			arChilds	=	getChilds(rsNewCat("CategoryID"))
			for ik=0 to rsNewCat("YoungestChildren")-1 
	%>
			<option value="<%=arChilds(0,ik)%>">&nbsp;&nbsp;-<%=arChilds(1,ik)%></option>
	<%
			next
		end if	
	rsNewCat.movenext
	Loop
	%>
</select>
<%
end sub
		
%>
<%
	Sub SelNV(nvName)
		strNhanvien="Select NhanVienID,Ho_Ten,BranchName From Nhanvien n JOIN Branch b ON n.BranchID = b.BranchID"
		dim rsNV
		set rsNV=Server.CreateObject("ADODB.Recordset")
		rsNV.open strNhanvien,Con,1 
		%>
		<select name="<%=nvName%>" size="1">
		<option value="0" selected="selected">Mời chọn</option>
		<%
			Do while not rsNV.eof
		%>
			<option value="<%=rsNV("NhanVienID")%>">
				<%=rsNV("Ho_Ten")%>&nbsp;-&nbsp;<%=rsNV("BranchName")%>
			</option>
		<%
			rsNV.movenext
			Loop
		%>
		</select>
<%
	end sub
%>
<%
	Sub SelectNhanVien(WName,WK,isMember,room,title)
    dNow = getDateServer()
		sql = "Select * From ChucDanh where {fn UCASE(ChucDanh.Description)}=N'"&UCASE(title)&"'"
		set rsNV=Server.CreateObject("ADODB.Recordset")
		rsNV.open sql,Con,1
		if not rsNV.eof then
			ChucdanhID	=	rsNV("ChucdanhID")
		end if	
		sql = "SELECT  Nhanvien.NhanVienID, Nhanvien.Ho_Ten, StaffContract.EndDate FROM  Nhanvien INNER JOIN StaffContract ON Nhanvien.NhanVienID = StaffContract.NhanVienID INNER JOIN PhongBan ON  StaffContract.PhongID =  PhongBan.PhongID "
		if isMember=0 then
			sql = sql + " and isMember='"& isMember &"' or EndDate=<'"& dNow &"'"
		elseif isMember=2 or isMember=3 or isMember=4 then
			sql = sql + " and isMember='"& isMember &"'"
		elseif isMember = 1 then
			sql = sql + " and isMember='"& isMember &"' and EndDate>='"& dNow &"' "
		elseif isMember = 5 then
			sql = sql + " and EndDate<'"& dNow &"'"
		elseif isMember = 6 then
			sql = sql + " and (isMember=1 or isMember=2 or isMember=3) and EndDate>='"& dNow &"'"
		end if
		
		if room <> "0" and room<>"" then
			sql = sql + " and {fn UCASE(PhongBan.Description)}=N'"&UCASE(room)&"'"
		end if
		if title <> "0" and title<>"" then
			sql = sql + " and ChucdanhID <= '"&ChucdanhID&"'"
		end if	
		sql=sql+" order by Ho_Ten,isMember,StaffContract.ID DESC "
		'Response.Write(sql)
'        0 Đã nghỉ việc
'        1 Đang hoạt động
'        2 Bán thời gian
'        3 Đối tác
'        4 Khác		
		set rsNV=Server.CreateObject("ADODB.Recordset")
		rsNV.open sql,Con,1 
		%>
		<select name="<%=WName%>" size="1">
		<option value="0" selected="selected">Mời chọn</option>
		<%
			Do while not rsNV.eof
		%>
		<option value="<%=rsNV("NhanVienID")%>" <%if WK=rsNV("NhanVienID") then%> selected="selected"<%end if%>>
			<%=rsNV("Ho_Ten")%>
		</option>
		<%
			rsNV.movenext
			Loop
		%>
		</select>
<%
	end sub
	
	Sub SelectTKThuChi(TKName,TKThuChiID,isThuChi)
		sql="Select * From TKThuChi where isThuChi="&isThuChi
		set rsTKThuChi=Server.CreateObject("ADODB.Recordset")
		rsTKThuChi.open sql,Con,1 
		%>
		<select name="<%=TKName%>" size="1">
		<option value="0" selected="selected">Mời chọn</option>
		<%
			Do while not rsTKThuChi.eof
			if TKThuChiID=rsTKThuChi("TKThuChiID") then
		%>
			<option value="<%=rsTKThuChi("TKThuChiID")%>" selected="selected">
				<%=rsTKThuChi("Mota")%>
			</option>
		<%
		else
		%>
			<option value="<%=rsTKThuChi("TKThuChiID")%>">
				<%=rsTKThuChi("Mota")%>
			</option>
		<%
			End if
			rsTKThuChi.movenext
			Loop
			set	rsTKThuChi = nothing
		%>
		</select>
<%
	end sub
	
	function GetTKThuChi(TKThuChiID)
		strtemp=""
		sql="Select * From TKThuChi where TKThuChiID="&TKThuChiID
		set rsTKThuChi=Server.CreateObject("ADODB.Recordset")
		rsTKThuChi.open sql,Con,1 
		if not rsTKThuChi.eof then
			strtemp = Trim(rsTKThuChi("MoTa"))
		end if
		GetTKThuChi	=	strtemp
	end function
	
	function QuyenThuChi(strChucDanh)
		iKT	=	0
		select case Uni2NONE(LCase(strChucDanh))
			case	"ban giam doc"
				iKT		=	1
			case	"ke toan"
				iKT		=	2
			case	"thu quy"
				iKT		=	3
		end select
		QuyenThuChi	=	iKT
	end function 
	
	Sub SelectTKTaiSan(TKName,TKTaiSanID,isLoaiTS)
		sql="Select * From TKTaiSan where isLoaiTS="&isLoaiTS
		set rsTemp=Server.CreateObject("ADODB.Recordset")
		rsTemp.open sql,Con,1 
		%>
		<select name="<%=TKName%>" size="1">
		<option value="0" selected="selected">Mời chọn</option>
		<%
			Do while not rsTemp.eof
			if TKTaiSanID=rsTemp("ID") then
		%>
			<option value="<%=rsTemp("ID")%>" selected="selected">
				<%=rsTemp("Mota")%>
			</option>
		<%
		else
		%>
			<option value="<%=rsTemp("ID")%>">
				<%=rsTemp("Mota")%>
			</option>
		<%
			End if
			rsTemp.movenext
			Loop
			set	rsTemp = nothing
		%>
		</select>
<%
	end sub
	function GetMaTS(ID)
		sql="Select * From TKTaiSan where ID="&ID
		set rsTemp=Server.CreateObject("ADODB.Recordset")
		rsTemp.open sql,Con,1 
		if not rsTemp.eof then
			GetMaTS	=	rsTemp("MoTa")
		else
			GetMaTS=""
		end if
	end function
	
	function GetKieuTS(ID)
		sql="Select * From TKTaiSan where ID="&ID
		set rsTemp=Server.CreateObject("ADODB.Recordset")
		rsTemp.open sql,Con,1 
		if not rsTemp.eof then
			GetKieuTS	=	rsTemp("isLoaiTS")
		else
			GetKieuTS=2
		end if
	end function	
	
function SelectMaHD(WName,WK)	
		strMaHD="SELECT inProductID,Maso,ProviderName,DateTime FROM inputProduct INNER JOIN Provider ON inputProduct.ProviderID = Provider.ProviderID"
		dim rsMaHD
		set rsMaHD=Server.CreateObject("ADODB.Recordset")
		rsMaHD.open strMaHD,Con,1 
		%>
		<select name="<%=WName%>" size="1">
		<option value="0" selected="selected">Mời chọn</option>
		<%
			Do while not rsMaHD.eof
				maso = rsMaHD("Maso")
				ProviderName= rsMaHD("ProviderName")
				Datet =rsMaHD("DateTime")
			%>
			<option value="<%=rsMaHD("inProductID")%>"<% if WK=rsMaHD("inProductID") then %>selected="selected"<%end if%>>
				<%=maso%> ( <%=ProviderName%> -- <%=Day(Datet)%>/<%=Month(Datet)%>/<%=year(Datet)%>)
			</option>
			<%
				rsMaHD.movenext
			loop
		%>
		</select>
		<%
			rsMaHD.close
			set rsMaHD = nothing
end function
	
function GetNameNV(MaNV)
		strNhanvien="Select Ho_Ten From Nhanvien where NhanVienID = '"& MaNV &"'"
		set rsNV=Server.CreateObject("ADODB.Recordset")
		rsNV.open strNhanvien,Con,1
		if not rsNV.eof then
			GetNameNV = rsNV("Ho_Ten")
		else
			GetNameNV = ""	
		end if
		set rsNV =nothing
end function

sub selLoaiSP(stt,catG)
%>
		<select name="selLSP<%=stt%>" size="1">
			<option value="0" selected="selected">Lựa chọn</option>
			<%
			strSel="Select CategoryID,CategoryName,YoungestChildren From NewsCategory where CategoryLoai=3 and categoryLevel = 1"
			dim rsNewCat
			set rsNewCat=Server.CreateObject("ADODB.Recordset")
			rsNewCat.open strSel,Con,1 
			%>
			<%
			
			Dim arChilds 
			Do while not rsNewCat.eof
			if rsNewCat("YoungestChildren") >0 then 'Nếu có con
			%>
				<option value="<%=rsNewCat("CategoryID")%>">*<%=rsNewCat("CategoryName" )%></option>
			<%	
				arChilds	=	getChilds(rsNewCat("CategoryID"))
				for ik=0 to rsNewCat("YoungestChildren")-1 
					If catG = arChilds(0,ik) Then
						%><option value="<%=arChilds(0,ik)%>"  selected="selected">&nbsp;&nbsp;-<%=arChilds(1,ik)%></option><%
					Else
					%>
						<option value="<%=arChilds(0,ik)%>" >&nbsp;&nbsp;-<%=arChilds(1,ik)%></option>
					<%
					End If
				Next	
			Else ' Nếu không có con
				If catG = rsNewCat("CategoryID") Then
				%>
					<option value="<%=rsNewCat("CategoryID")%>" selected="selected">*<%=rsNewCat("CategoryName" )%></option>
				<%
				Else
				%>
					<option value="<%=rsNewCat("CategoryID")%>">*<%=rsNewCat("CategoryName" )%></option>
				<%	
				End If
			End If
				rsNewCat.movenext
				Loop
			%>
	</select>
<%
end sub
		
%>

<%
sub selLSP(stt)
%>
	<select name="selLSP<%=stt%>" size="1">
	<option value="0" selected="selected">Lựa chọn</option>
	<%
	strSel="Select CategoryID,CategoryName,YoungestChildren From NewsCategory where CategoryLoai=3 and categoryLevel = 1"
	dim rsNewCat
	set rsNewCat=Server.CreateObject("ADODB.Recordset")
	rsNewCat.open strSel,Con,1 
	%>
	<%
	Dim arChilds 
	Do while not rsNewCat.eof
	%>
		<option value="<%=rsNewCat("CategoryID")%>">*<%=rsNewCat("CategoryName")%></option>
	<%
		if rsNewCat("YoungestChildren") >0 then
			arChilds	=	getChilds(rsNewCat("CategoryID"))
			for ik=0 to rsNewCat("YoungestChildren")-1 
	%>
			<option value="<%=arChilds(0,ik)%>">&nbsp;&nbsp;-<%=arChilds(1,ik)%></option>
	<%
			next
		end if	
	rsNewCat.movenext
	Loop
	%>
	</select>
<%
end sub

function getNhanVienFromID(NhanVienID)
	strNV="SELECT Ho_Ten FROM Nhanvien WHERE NhanVienID ='"&NhanVienID&"'"
	set reNV=Server.CreateObject("ADODB.Recordset")
	reNV.open strNV,Con,1
	if not reNV.eof then
		getNhanVienFromID	=	reNV("Ho_Ten")
	else
		getNhanVienFromID  =  ""
	end if
	reNV.close()
	set reNV = nothing
end function

function GetPhiVanChuyen(SanPhamUser_ID)
	strPay 	= "Select Phivanchuyen From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		GetPhiVanChuyen	=	rsPay("Phivanchuyen")
	else
		GetPhiVanChuyen  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetThuKhac(SanPhamUser_ID)
	strPay 	= "Select ThuKhac From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		GetThuKhac	=	rsPay("ThuKhac")
	else
		GetThuKhac  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetChiKhac(SanPhamUser_ID)
	strPay 	= "Select ChiKhac From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		GetChiKhac	=	rsPay("ChiKhac")
	else
		GetChiKhac  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetGhiChu(SanPhamUser_ID)
	strPay 	= "Select GhiChu From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		GetGhiChu	=	rsPay("GhiChu")
	else
		GetGhiChu  =  ""
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetCuocBuuDienID(SanPhamUser_ID)
	strPay 	= "Select Cuocbuudien From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		if isnumeric(rsPay("Cuocbuudien")) =  true then
			GetCuocBuuDienID	=	Clng(rsPay("Cuocbuudien"))
		else
			GetCuocBuuDienID  =  0
		end if
	else
		GetCuocBuuDienID  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

function fTotalSaleOff(SanPhamUser_ID)
	sql = "SELECT TruTrongTaiKhoan FROM SanPham_Pay where SanPhamUser_ID='"& SanPhamUser_ID &"'"
	Set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,con,1
	iTemp	=	0
	if not rsTemp.eof then
		iTemp 	=	GetNumeric(rsTemp("TruTrongTaiKhoan"),0)
	end if
	rsTemp.close
	fTotalSaleOff	=	iTemp
end function


function GetCuocBuuDienThucID(SanPhamUser_ID)
	strPay 	= "Select CuocbuudienThuc From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		if isnumeric(rsPay("CuocbuudienThuc")) =  true then
			GetCuocBuuDienThucID	=	Clng(rsPay("CuocbuudienThuc"))
		else
			GetCuocBuuDienThucID  =  0
		end if
	else
		GetCuocBuuDienThucID  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetMaBuuDien(SanPhamUser_ID)
	strPay 	= "Select MaBuuDien From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		GetMaBuuDien	=	Trim(rsPay("MaBuuDien"))
	else
		GetMaBuuDien  =  ""
	end if
	rsPay.close()
	set rsPay = nothing
end function

function GetKhoiLuongThucID(SanPhamUser_ID)
	strPay 	= "Select TLThuc From SanPham_pay where SanPhamUser_ID = '"& SanPhamUser_ID &"'"
	set rsPay=Server.CreateObject("ADODB.Recordset")
	rsPay.open strPay,Con,1
	if not rsPay.eof then
		if isnumeric(rsPay("TLThuc")) =  true then
			GetKhoiLuongThucID	=	Clng(rsPay("TLThuc"))
		else
			GetKhoiLuongThucID  =  0
		end if
	else
		GetKhoiLuongThucID  =  0
	end if
	rsPay.close()
	set rsPay = nothing
end function

Function GetInfoSanPhamNhap(iItemID)
	Dim aParameters 
	sqlNews="SELECT * from SanPhamNhap where NewsId=" & iItemID
	Set rsNews=Server.CreateObject("ADODB.Recordset")
	rsNews.open sqlNews,con,3
	if rsNews.eof then
		rsNews.close
		set rsNews=nothing
		aParameters = Array("","","","","","","")
		exit Function
	else
		idsanpham   =   rsNews("idsanpham")
		Title		=	rsNews("Title")
		nxb			=	rsNews("nxb")
		tacgia		=	rsNews("tacgia")
		namxuatban	=	rsNews("namxuatban")
		trongluong	=	rsNews("trongluong")
		'                      0        1      2    3     4            5        6
		aParameters = Array(idsanpham,Title,tacgia,nxb,trongluong,namxuatban)
		rsNews.close
		set rsNews=nothing
	end if
GetInfoSanPhamNhap = aParameters
End Function

		
function GetNumInvoiceReturnProvice(ProductID)
	sql	=	"SELECT  SUM(SLTraNCC) AS iSum FROM TraSach WHERE ProductID = "&ProductID
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetNumInvoiceReturnProvice = GetNumeric(rsTemp("iSum"),0)
	else
		GetNumInvoiceReturnProvice=0
	end if
	set rsTemp = nothing
end function

function GetStaffReturnStore(ProductID)
	sql	=	"SELECT  NhanVienID FROM TraSach WHERE ProductID = "&ProductID
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		Employees 	=	getNhanVienFromID(rsTemp("NhanVienID"))
	else
		Employees	=	""
	end if
	set rsTemp = nothing
	GetStaffReturnStore	=	Employees
end function

function GetDateReturnStore(ProductID)
	sql	=	"SELECT  NgayTra FROM TraSach WHERE ProductID = "&ProductID
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetDateReturnStore 	=	rsTemp("NgayTra")
	else
		GetDateReturnStore	=	null
	end if
	set rsTemp = nothing
end function

function GetNumInvoiceOutStore(ProductID)
	sql = "SELECT SUM(SoLuong) AS iCount FROM XuatKho WHERE ProductID = '"&ProductID&"'"
	set rsOutStore =  Server.CreateObject("ADODB.recordset")
	rsOutStore.open sql,Con,1
	if not rsOutStore.eof then
		if	isnumeric(rsOutStore("iCount")) = true then
			GetNumInvoiceOutStore = rsOutStore("iCount")
		else
			GetNumInvoiceOutStore=0
		end if
	else
		GetNumInvoiceOutStore=0
	end if
	set rsOutStore = nothing
end function

function GetNumProductNhapKho(ProductID)
	sql = "SELECT Number FROM Product WHERE ProductID = '"&ProductID&"'"
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetNumProductNhapKho = rsTemp("Number")
	else
		GetNumProductNhapKho=0
	end if
	set rsTemp = nothing
end function

function GetTotalXuatKho(SanPham_User_ID)
	sql = "SELECT SUM(SoLuong) AS iCount FROM XuatKho WHERE SanPham_User_ID = '"&SanPham_User_ID&"'"
	set rs1 =  Server.CreateObject("ADODB.recordset")
	rs1.open sql,Con,1
	if not rs1.eof then
		GetTotalXuatKho = GetNumeric(rs1("iCount"),0)
	else
		GetTotalXuatKho=0
	end if
	set rs1 = nothing
end function

function GiaNhapKho(SanPham_User_ID)
	sql =       "SELECT Product.ProductID,NewsID,Number,Giabia,Price,VAT"
	sql = sql + " FROM Product INNER JOIN XuatKho ON Product.ProductID = XuatKho.ProductID"
	sql = sql + " where SanPham_User_ID ="&SanPham_User_ID
	set rsTemp = Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1 
'	do while not rstemp.eof 
	
	if not rsTemp.eof then
		Price  	=	rsTemp("Price")
		iVAT	=	rsTemp("VAT")	
		Price = Price + Price*iVAT/100
		GiaNhapKho = Price
	else
		GiaNhapKho = 0
	end if
		'Response.Write("Giá "&Price&"<br>")
'		rsTemp.movenext
'	loop
	set rsTemp = nothing
end function

function TongTrongLuong(User_ID)
	Set rsUser=server.CreateObject("ADODB.Recordset")
	sqlfunc="SELECT	SanPham_ID " &_
		"FROM  SanPham_User " &_
		"WHERE     (SanPhamUser_ID = " & User_ID & ") and re_newsid = 0"	
	rsUser.open sqlfunc,con,1
	iTotalKL	=	0
	Do while not rsUser.EOF
		NewsID = rsUser("SanPham_ID")
		sqlfunc = "Select Trongluong from SanPhamNhap where NewsID='"& NewsID &"'"
		Set rss=server.CreateObject("ADODB.Recordset")
		rss.open sqlfunc,con,1
		if not rss.eof then
			iTotalKL	=	 iTotalKL + rss("Trongluong")
		end if
		set rss = nothing
		rsUser.MoveNext
	Loop
	rsUser.close
	TongTrongLuong = iTotalKL
	
end function

function TongTienTrenDonHang(User_ID,strCMND)
	Set rsUser=server.CreateObject("ADODB.Recordset")
	sqlfunc="SELECT	* " &_
		"FROM  SanPham_User " &_
		"WHERE     (SanPhamUser_ID = " & User_ID & ") and re_newsid = 0"	
	rsUser.open sqlfunc,con,1
	sTotal=0
	iTotalSP	=	1
	Do while not rsUser.EOF
		Gia_sp=rsUser("Sanpham_Gia")
		Soluong=rsUser("SanPham_Soluong")
		on error Resume next
		Gia=CLng(Gia_sp)*CLng(Soluong)
		sTotal = sTotal + Gia
		rsUser.MoveNext
	Loop
	iSaleOff = fTotalSaleOff(User_ID)
	tTotal= sTotal+ GetPhiVanChuyen(User_ID)+GetCuocBuuDienID(User_ID) + GetThuKhac(User_ID) - iSaleOff - GetTienKM(User_ID)
	rsUser.close
	TongTienTrenDonHang = tTotal	
end function

function isNhapKho(SanPham_User_ID)
	strTempNK = "Select * from XuatKho where SanPham_User_ID = '"& SanPham_User_ID &"'"
	Set rsTempNK=server.CreateObject("ADODB.Recordset")
	rsTempNK.open strTempNK,con,1
	if  not rsTempNK.eof then
		isNhapKho = true
	else
		isNhapKho = false
	end if
	set rsTempNK = nothing
end function

function ChuKy(ID)
	iCk = false
	sql =	"select * From PhieuKeToan where ID="&ID
	set rsCK = Server.CreateObject("ADODB.recordset")
	rsCK.open sql,Con,1	
	if not rsCK.eof then
		if rsCK("ChukyKT")<>0 and rsCK("ChukyTQ")<>0 and rsCK("dongy")<>0 then
			iCk	=	true
		end if
	end if
	ChuKy	=	iCk
	set rsCK	=	nothing
end function 

function SoLuongSPTrongDonH(SanPham_User_ID)
	sql = "SELECT   SanPham_Soluong FROM SanPham_User WHERE  re_newsid = 0 and SanPham_User_ID = "&SanPham_User_ID
	set rsTemp = Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	
	if not rsTemp.eof then
		if isnumeric(rsTemp("SanPham_SoLuong")) = true then
			SoLuongSPTrongDonH	=	rsTemp("SanPham_SoLuong")
		else
			SoLuongSPTrongDonH= 0
		end if		
	end if
	set rsTemp = nothing
end function
Function SoLuongSPNhapKho(ProductID)
	sql = "Select Number From Product where ProductID ="&ProductID
	set rsTemp = Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		SoLuongSPNhapKho	=	rsTemp("Number")
	end if
	set rsTemp = nothing
end function

function getInt(Doituong)
	if Doituong <> "" and isnumeric(Doituong) = true then
		getInt = Clng(Doituong)
	else
		getInt = 0
	end if
end function

function GetTotalSPinHD(inProductID)
	sqlTemp =  "SELECT  SUM(Product.Number) AS TotalSP FROM inputProduct INNER JOIN Product ON inputProduct.inProductID = Product.inProductID WHERE inputProduct.inProductID = '"& inProductID &"'"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sqlTemp,Con,1
	if not rsTemp.eof then
		GetTotalSPinHD = rsTemp("TotalSP")
	else
		GetTotalSPinHD = 0
	end if
	rsTemp.close
	set rsTemp = nothing
end function

function GetTTien(inProductID)
	sqlTemp = "SELECT  Number, Price, VAT FROM Product WHERE inProductID = '"& inProductID &"'"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	TTien = 0
	rsTemp.open sqlTemp,Con,1	
	do while not rsTemp.eof
		iNum 	=	rsTemp("Number")
		if isnumeric(rsTemp("Price")) = true then
			iPrice 	=	CDbl(rsTemp("Price"))
		else
			iPrice = 0
		end if
		iVAT 	=	rsTemp("VAT")
		Tien = Tien + Tien*iVAT/100
		Tien = iPrice*iNum
		TTien = TTien + Tien		
		rsTemp.movenext
	loop
	rsTemp.close
	set rsTemp = nothing
	GetTTien = TTien
end function

function GetNewsIDFromIDSanPhamNhap(IDSanPham)
	sql = "Select NewsID From SanPhamNhap where idsanpham = '"&IDSanPham&"'"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,Con,1	
	if not rsTemp.eof then
		GetNewsIDFromIDSanPhamNhap = rsTemp("NewsID")
	else
		GetNewsIDFromIDSanPhamNhap = ""
	end if
	rsTemp.close
	set rsTemp = nothing
end function


function isXuatKho(SanPhamUser_ID)
	sql = "SELECT NewsID,SanPham_User_ID,SanPham_Soluong  "
	sql= sql + " FROM SanPham_User INNER JOIN SanPhamNhap ON SanPham_User.SanPham_ID = SanPhamNhap.NewsID "
	sql= sql + " WHERE SanPhamUser_ID = '"& SanPhamUser_ID &"' and re_newsid = 0"
	set rsTemp123=Server.CreateObject("ADODB.Recordset")
	rsTemp123.open sql,Con,1 
	Do while not rsTemp123.eof
		SanPham_Soluong	=	rsTemp123("SanPham_Soluong")
		inum_order		= 	inum_order + SanPham_Soluong
		iout			=	GetTotalXuatKho(rsTemp123("SanPham_User_ID"))
		NewsID			=	rsTemp123("NewsID")
		istoretemp		= 	GetNumInventoryGoods(NewsID)
		'Response.Write("istoretemp"&istoretemp&"SanPham_Soluong"&SanPham_Soluong)
		if istoretemp >	SanPham_Soluong then
			istoretemp	= SanPham_Soluong
		end if
		istore_order	=	 istoretemp + istore_order
		inum_out	=	inum_out	+ iout		
		rsTemp123.movenext
	loop
	set rsTemp123 = nothing
	if inum_order = inum_out then
		isXuatKho	= true
	else
		isXuatKho	= false
	end if
	'Response.Write("inum_out"&inum_out)
end function 

function getNVNoCu(NhanVienID)
	iNocu	=	0
	sql = "select Nocu from NVCongNo where NhanVienID="&NhanVienID
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof  then
		iNocu	=	GetNumeric(rsTemp("Nocu"),0)
	end if
	getNVNoCu	=	iNocu	
end function

function GetTienKM(User_ID)
	GiaKM = 0
	sql = "SELECT Gia FROM SanPhamUser INNER JOIN SPKhuyenMai ON SanPhamUser.IDKhuyenMai = SPKhuyenMai.ID where SanPhamUser_ID='"& User_ID &"' and status = 1"
	set rskm1	=	Server.CreateObject("ADODB.recordset")
	rskm1.open sql,con,1
	if not rskm1.eof then	
		GiaKM	=	rskm1("Gia")
	end if
	GetTienKM	=	GiaKM
end function

function GetIDNhanVienUserName(UserName)
	sql = "Select * from [User] where UserName = '"& UserName &"'"
 	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	IDNV = 0
	rsTemp.open sql,con,1
	if not rsTemp.eof then	
	IDNV	=	rsTemp("IDNhanVien")
	end if
	GetIDNhanVienUserName	=	IDNV
end function

function GetUserName(IDNhanVien)
	sql = "Select * from [User] where IDNhanVien = '"& IDNhanVien &"'"
 	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	UserName = 0
	rsTemp.open sql,con,1
	if not rsTemp.eof then	
	UserName	=	rsTemp("UserName")
	end if
	GetUserName	=	UserName
end function
%>

<%
sub GetEChuKy(NhanVienID,iWidth,iHeigth,iEdit)
	sql = "Select * from Nhanvien where NhanVienID = '"& NhanVienID &"'"
 	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	eChuKy = ""
	rsTemp.open sql,con,1
	if not rsTemp.eof then	
	eChuKy	=	rsTemp("eChuKy")
	if iWidth > 0 then
		str	= "width="&iWidth
	elseif iHeigth > 0 then
		str	= "height="&iHeigth
	end if
%>	
<img src="<%=eChuKy%>"  <%=str%> border="0" align="absmiddle" alt="<%=Ho_Ten%>" />
<%	end if		
end sub
%>

<%
sub ImgDimension(img) 
   dim myImg, fs 
   Set fs= CreateObject("Scripting.FileSystemObject") 
   if not fs.fileExists(img) then exit sub 
   set myImg = loadpicture(img) 
   iWidth = round(myImg.width / 26.4583) 
   iHeight = round(myImg.height / 26.4583) 
   iType = myImg.Type 
   select case iType 
   case 0 
       iType = "None" 
   case 1 
       iType = "Bitmap" 
   case 2 
       iType = "Metafile" 
   case 3 
       iType = "Icon" 
   case 4 
       iType = "Win32-enhanced metafile" 
   end select 
   set myImg = nothing 
end sub 
%> 

<%
' hàm số hóa đơn chứng từ
function fSoHDCT(ID,DateTime,iLoaiHDCT)
	strHDCT	= ""
	Select case iLoaiHDCT
		case 1 ' Hóa đơn
			strHDCT	=	"HĐ"
		case 2 ' Chứng từ
			strHDCT	=	"CT"
		case 3 ' Tài sản
			strHDCT	=	"TS"
		case 4 ' Vay vốn
			strHDCT	=	"VV"
		case 5 ' Tiền mặt
			strHDCT	=	"TM"
		case 6 ' Hóa đơn Thu
			strHDCT	=	"HĐT"
		case 7 ' Hóa đơn chi
			strHDCT	=	"HĐC"			
		case 8 ' Hóa đơn tạm ứng
			strHDCT	=	"TU"
		case 9 ' Hóa nộp tiền mặt
			strHDCT	=	"NT"
		case 10 ' trả lương
			strHDCT	=	"TL"						
		case 11 ' trả lương
			strHDCT	=	"GDT"			
	end Select
	
	CSoHD	=	strHDCT&ID&"/"&Month(DateTime)&right(DateTime,2)
	fSoHDCT	=	CSoHD
end function

function fIDFormSoHDCT(CSoHD,DateTime,iLoaiHDCT)
	strHDCT	= ""
	Select case iLoaiHDCT
		case 1 ' Hóa đơn
			strHDCT	=	"HĐ"
		case 2 ' Chứng từ
			strHDCT	=	"CT"
		case 3 ' Tài sản
			strHDCT	=	"TS"
		case 4 ' Vay vốn
			strHDCT	=	"VV"
		case 5 ' Tiền mặt
			strHDCT	=	"TM"
		case 6 ' Hóa đơn Thu
			strHDCT	=	"HĐT"
		case 7 ' Hóa đơn chi
			strHDCT	=	"HĐC"			
		case 8 ' Hóa đơn tạm ứng
			strHDCT	=	"TU"
		case 9 ' Hóa nộp tiền mặt
			strHDCT	=	"NT"
		case 10 ' trả lương
			strHDCT	=	"TL"									
	end Select
	CSoHD	=	replace(CSoHD,strHDCT,"")
	i = InStr(CSoHD, "/")
	CSoHD = Mid(CSoHD,1, i - 1)	
	ID	=	GetNumeric(CSoHD,0)
	if ID <> 0 then
		fIDFormSoHDCT	=	ID
	else
		fIDFormSoHDCT	=	""
	end if
end function
%>

<%
function isArrExits(aMang,phanTu)
	isTest	=	false
	for kt=0 to Ubound(aMang)
		if aMang(kt) = phanTu then
			isTest	=	True
			exit for
		end if
	next
	isArrExits	=	isTest
end function

%>
<%
function Get_IDGroup(name_group)
	sql = "select * FROM EmailNhom where TenNhom=N'"& name_group &"'"
	set rsIDGroup	=	Server.CreateObject("ADODB.recordset")
	rsIDGroup.open sql,con,1
	id_group	=	0
	if not rsIDGroup.eof then
		id_group	=	rsIDGroup("IDNhomEmail")
	end if
	Get_IDGroup	=	id_group
end function



function Get_IDGroup_FromEmail(Email,m_group)
	' m_group = "IDXungHo"
	' m_group =	"IDTamly"
	' m_group = "IDCongViec"
	sql = "SELECT "& m_group &" FROM Email where Email='"& Email &"'"
 	set rs_group	=	Server.CreateObject("ADODB.recordset")
	IDGroup = 0
	rs_group.open sql,con,1
	if not rs_group.eof then	
		IDGroup	=	rs_group(m_group)
	end if
	Get_IDGroup_FromEmail	=	IDGroup	
	set rs_group	=	nothing
end function

function isCheckExpress(GiaoHang_times)

	if InStr(GiaoHang_times," 24 giờ")>0 or InStr(GiaoHang_times," 6 giờ")>0  or InStr(GiaoHang_times," 5 giờ")>0 or InStr(GiaoHang_times,"Từ 1 đến 2 ngày trừ ngày lễ và chủ nhật")>0 or InStr(GiaoHang_times,"Giao nhanh")>0 then
		isCheckExpress  =  true
	else
		isCheckExpress 	=	false
	end if
	
end function 

function isHanoiExpress(Diachi)
	Diachi	=	UCASE(Diachi)
	if InStr(Diachi,"HA NOI")>0 or InStr(Diachi,"HÀ NỘI")>0  or InStr(Diachi,"HANOI")>0 then
		isHanoiExpress  =  true
	else
		isHanoiExpress 	=	false
	end if	
end function
%>

<%
function isExistEmail(strEmail)
	strEmail	=	trim(strEmail)
	iTontai	= 0
	sql 	=	"Select * from Email where Email=N'"&strEmail&"'"
	set rsTemp	=	server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,con,1
	if not rsTemp.eof then
		iTontai	=	1
	else
		iTontai =	0
	end if	
	isExistEmail	=	iTonTai
	set rsTemp	= nothing
end function 
%>
<%
function Get_Name_Group(ID)
	m_name 	=	""
	sql 	=	"Select * from EmailNhom where IDNhomEmail ='"&ID&"'"
	set rsTemp	=	server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,con,1
	if not rsTemp.eof then
		m_name	=	rsTemp("TenNhom")
	end if
	Get_Name_Group	=	m_name
end function

function GetRootCategory(NewsID)
	m_name 	=	""
	sql = "SELECT top 1 c.CategoryName,c.ParentCategoryID FROM News AS n INNER JOIN NewsDistribution AS d ON n.NewsID = d.NewsID INNER JOIN NewsCategory AS c ON d.CategoryID = c.CategoryID  where n.NewsID="&NewsID
	set rsTemp	=	server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,con,1
	if not rsTemp.eof then
		ParentCategoryID	=	rsTemp("ParentCategoryID")
		if ParentCategoryID = 0 then
			m_name	=	rsTemp("CategoryName")
		end if
		i = 5
		do while ParentCategoryID <> 0 
			sql = "Select top 1 CategoryName,ParentCategoryID from NewsCategory where CategoryID="&ParentCategoryID
			set rsTemp1	=	server.CreateObject("ADODB.Recordset")
			rsTemp1.open sql,con,1
			if not rsTemp1.eof then
				ParentCategoryID = rsTemp1("ParentCategoryID")
				m_name	=	rsTemp1("CategoryName")
			end if
			set rsTemp1 = nothing
			i = i + 1
			if i > 5 then
				exit do
			end if
		loop
	end if
	set rsTemp	=	nothing
	
	GetRootCategory	=	m_name	
end function

function GetDetailInput(NewsID,isReturn,iShowTable)
' isReturn = 0 Trả về ngày nhập gần nhất
' isReturn = 1 Tổng giá nhập gần nhất
' isReturn = 2 Tổng giá bìa
' isReturn = 4 Các ngày nhập của sách tồn
' iShowTable= 0 không hiện thị bảng
' iShowTable= 1 Hiện thị bảng
' iShowTable= 2 Hiển thị ngày

	sqlTemp = "SELECT SanPhamNhap.NewsID, Product.Number, Product.ProductID,Product.Giabia, Product.Price, inputProduct.inProductID, inputProduct.Maso, inputProduct.DateTime FROM SanPhamNhap INNER JOIN Product ON SanPhamNhap.NewsID = Product.NewsID INNER JOIN inputProduct ON Product.inProductID = inputProduct.inProductID WHERE (SanPhamNhap.NewsID = '"&NewsID&"') "
	sqlTemp=sqlTemp+" and inputProduct.AccountingSigna<>0 and inputProduct.StoreSigna<>0 and inputProduct.CreaterSigna<>0 "
	sqlTemp=sqlTemp+" ORDER BY inputProduct.DateTime DESC"
	set rsTemp1	=	server.CreateObject("ADODB.Recordset")
	rsTemp1.open sqltemp,con,1
	if not rsTemp1.eof  and iShowTable = 1 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Ngày nhập</strong> </div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>SL</strong></div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Bìa</strong></div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Nhập</strong></div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Tồn</strong></div></td>
  </tr>
<%	
	end if
	fConverTotal 	= 	0
	fPriceTotal		= 	0
	do while not rsTemp1.eof
		if GetNumInvoiceOutStore(rsTemp1("ProductID")) < rsTemp1("Number") then
			fConverTotal	= fConverTotal+rsTemp1("Giabia")
			fPriceTotal	=	fPriceTotal+rsTemp1("Price")
			if isReturn	=	4 and iShowTable = 2 then
%>
<a href="../thongke/Report_SoHD.asp?inProductID=<%=rsTemp1("inProductID")%>" target="_blank"> <%=GetFullDate(rsTemp1("DateTime"))%></a><br />
<%							
			end if
		end if
		if isReturn	=	0 then
			GetDetailInput = rsTemp1("DateTime")
			exit do
		end if

		if  iShowTable = 1 then
%>	
  <tr>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><a href="../thongke/Report_SoHD.asp?inProductID=<%=rsTemp1("inProductID")%>" target="_blank"> <%=GetFullDate(rsTemp1("DateTime"))%></a></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsTemp1("Number")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsTemp1("Giabia"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsTemp1("Price"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsTemp1("Number")-GetNumInvoiceOutStore(rsTemp1("ProductID"))%></td>
  </tr>

<%
		end if
		rsTemp1.movenext
	loop
	set rsTemp1	=	nothing
%>
<%if  iShowTable = 1 then%>
</table>
<%end if%>
<%
	if isReturn	=	1 then
		GetDetailInput	=	fPriceTotal
	elseif isReturn	=	2 then
		GetDetailInput	= 	fConverTotal
	end if
end function
%>
<%
function GetNumInventore(ProductID)
	iNumOutStore  = GetNumInvoiceOutStore(ProductID)
	iNumInStore	  =	rsProd("Number")
	iNumReturnStore=GetNumInvoiceReturnProvice(ProductID)
	iNumInventory= iNumInStore - iNumOutStore - iNumReturnStore
	iOutTotal	=	iOutTotal+iNumOutStore
	iInTotal	=	iInTotal+iNumInStore
	iReturnTotal	=	iReturnTotal+iNumReturnStore
	iInventoryTotal	=	iInventoryTotal +iNumInventory
end function

function GetDetailOutput(NewsID,isReturn,iShowTable)
' isReturn = 0 trả về ngày nhập gần nhất
' isReturn = 1 Tổng giá nhập gần nhất
' isReturn = 2 Tổng giá bìa
' iShowTable= 0 không hiện thị bảng
	sqlTemp = "SELECT SELECT  XuatKho.SoLuong, SanPham_User.SanPham_Gia, SanPhamUser.NgayXuLy, SanPhamUser.OutStoreDate FROM Product INNER JOIN XuatKho ON Product.ProductID = XuatKho.ProductID INNER JOIN SanPham_User ON XuatKho.SanPham_User_ID = SanPham_User.SanPham_User_ID INNER JOIN SanPhamUser ON SanPham_User.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID WHERE (Product.NewsID = '"&NewsID&"') ORDER BY SanPhamUser.DateTime DESC"
	set rsTemp1	=	server.CreateObject("ADODB.Recordset")
	rsTemp1.open sqltemp,con,1
	Response.Write(sqlTemp1)
	if not rsTemp1.eof  and iShowTable = 1 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Ngày xuất </strong></div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>SL</strong></div></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><div align="center"><strong>Xuất</strong></div></td>
  </tr>
  <%	
	end if
	fConverTotal 	= 	0
	fPriceTotal		= 	0
	GetDate			= null
	do while not rsTemp1.eof
		fConverTotal	= fConverTotal+rsTemp1("Giabia")
		fPriceTotal	=	fPriceTotal+rsTemp1("Price")
		if isReturn	=	0 then
			GetDetailInput = rsTemp1("DateTime")
			exit do
		end if
%>
  <tr>
    <td style="<%=setStyleBorder(0,1,0,1)%>"><%=GetFullDate(rsTemp1("NgayXuLy"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=rsTemp1("SoLuong")%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=Dis_str_money(rsTemp1("SanPham_Gia"))%></td>
  </tr>
  <%
		rsTemp1.movenext
	loop
%>
</table>
<%
end function
%>

<script language="javascript">
function Outlogin()
{
<%
	sql = "update Nhanvien set Ngaycap='"& dNow &"'  where NhanVienID='"& Session("NhanVienID") &"'"
	set rsTemp=Server.CreateObject("ADODB.Recordset")
'	rsTemp.open sql,Con,1	
	set rsTemp = nothing
%>
}
</script>