<%  
'	XSOFT
'  	(C) Copyright XSOFT Corp. 2007
'  	**************************
' 	Nhom tu van thiet ke va phat trien phan mem  
'  	Quan ly nhan su, ban hang, ton kho, tai chinh ke toan, tai chinh gia dinh.
'  	Thiet ke website, thiet ke logo, catalog.
'  	website:www.xsoft.com.vn
'  	email:info@xsoft.com.vn – DT:04.2922.446
%>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_DateTime.asp" -->
<%
function fTotalCMTGioiThieu(CMND)
	sql	= "SELECT nguoi_gioi_thieu from Account where nguoi_gioi_thieu ='"& CMND &"'"
	Set rsCMT=Server.CreateObject("ADODB.Recordset")
	rsCMT.open sql,con,1
	totalCMT=0
	do while not rsCMT.eof
		totalCMT = totalCMT + 1
		rsCMT.movenext
	loop
	rsCMT.close
	fTotalCMTGioiThieu	=	totalCMT
end function

function fTotalIdea(CMND)
	sql = "SELECT CMND from Y_KIEN where CMND='"&CMND&"' and show=1"
	Set rsIdea=Server.CreateObject("ADODB.Recordset")
	iTotalIdea=0
	on error  Resume Next
	rsIdea.open sql,con,1
	do while not rsIdea.eof
		iTotalIdea = iTotalIdea + 1
		rsIdea.movenext
	loop
	rsIdea.close
	fTotalIdea	=	iTotalIdea
end function

function fTotalUser(CMND)
	sql = "SELECT SUM(TruTrongTaiKhoan) AS fTK FROM SanPham_pay INNER JOIN SanPhamUser ON SanPham_pay.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID where CMND='"& CMND &"' and SanPhamUser_Status <= 2"
	Set rsUsers=Server.CreateObject("ADODB.Recordset")
	rsUsers.open sql,con,1
	TotalUser	=	0
	if not rsUsers.eof then
		TotalUser 	=	GetNumeric(rsUsers("fTK"),0)
	end if
	rsUsers.close
	fTotalUser	=	TotalUser
end function



function fIniTKhoan(CMND)
	sql = "SELECT SUM(iniTK) AS fIniTK FROM TaiKhoan where CMND='"& CMND &"'"
	Set rsUsers=Server.CreateObject("ADODB.Recordset")
	rsUsers.open sql,con,1
	TotalIni	=	0
	if not rsUsers.eof then
		TotalIni 	=	GetNumeric(rsUsers("fIniTK"),0)
	end if
	rsUsers.close
	fIniTKhoan	=	TotalIni
end function

Function LamTronTien(fSoTien)
	if isnumeric(fSoTien) = false or fSoTien = "" then
		fSoTien = 0
	end if
	Temp 	= Cstr(fSoTien)
	if len(Temp) > 3 then
		Temp	= Right(Temp,3)
	end if
	iTemp 	= 	Cint(Temp)
	if iTemp > 500 and iTemp < 1000 then
		fTemp	=	1000 - iTemp
	elseif (Temp < 500) and Temp > 0 then
		fTemp	=	500 - iTemp
	elseif (iTemp = 500) or iTemp = 0 then
		fTemp = 0
	end if
	tTotal		=	fSoTien + fTemp
	
	LamTronTien = tTotal
end function

%>

<%Function GetNumeric(sValue,DefaultValue)
	Dim intValue
	if not IsNumeric(sValue) or trim(sValue)="" then
		intValue=DefaultValue
	else
		intValue=Clng(sValue)
	end if
	GetNumeric=intValue
End Function%>