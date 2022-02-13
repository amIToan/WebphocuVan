<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_human")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Dim Upload 'Su dung AspUpload
	Dim sError
	sError="" 'Luu ca'c thong ba'o loi
	Set Upload = Server.CreateObject("Persits.Upload")
	sql	= ""
	Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save
	iStatus	=	Trim(Upload.Form("iStatus"))
	Catup	=	Trim(Upload.Form("Catup"))	
if Catup = "staff" then
	Ho_Ten	=	Trim(Upload.Form("Ho_Ten"))
	CMT		=	Trim(Upload.Form("CMT"))
	salecode		=	Trim(Upload.Form("salecode"))
	NgaySinh=	Trim(Upload.Form("NgaySinh"))
	ThangSinh=	trim(UpLoad.Form("ThangSinh"))
	NamSinh	=	Trim(UpLoad.Form("NamSinh"))
	NgaySinh	=	ThangSinh&"/"&NgaySinh&"/"&NamSinh
	if isdate(NgaySinh) = true then
		NgaySinh	=FormatDatetime(NgaySinh)
	else
		NgaySinh = Null
	end if	
	
	NgayCap	=	Trim(Upload.Form("NgayCap"))
	ThangCap=	trim(UpLoad.Form("ThangCap"))
	NamCap	=	Trim(UpLoad.Form("NamCap"))
	NgayCap	=	ThangCap&"/"&NgayCap&"/"&NamCap
	if isdate(NgayCap) = true then
		NgayCap	=FormatDatetime(NgayCap)
	else
		NgayCap = Null
	end if		
	
	Noicap	=	Trim(Upload.Form("Noicap"))
	Hocvan	=	Trim(Upload.Form("Hocvan"))
	DanToc	=	Trim(Upload.Form("DanToc"))
	Tel		=	Trim(Upload.Form("Tel"))
	Mobile	=	Trim(Upload.Form("Mobile"))
	Email	=	Trim(Upload.Form("Email"))
	password_partner	=	Trim(Upload.Form("password_partner"))
	Cu_tru	=	Trim(Upload.Form("Cu_tru"))		
	Cu_tru=	Replace(Cu_tru,"<br>",chr(13) & chr(10))
	
	Diachi	=	Trim(Upload.Form("Diachi"))
	Diachi	=	Replace(Diachi,"<br>",chr(13) & chr(10))
	
	set imgNVfile = Upload.Files("imgNVfile")
	If imgNVfile Is Nothing Then
		ImgNV=""
	else
	   Filetype = Right(imgNVfile.Filename,len(imgNVfile.Filename)-Instr(imgNVfile.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" then
			sError=sError & "&nbsp;-&nbsp; Ảnh cán bộ không phải dạng ảnh<br>"
	   else
			imagename=CMT& "." & Filetype			
			imgNVfile.SaveAs server.MapPath("/administrator/images/Cky/")& "\" & imagename
			ImgNV	=	"/administrator/images/Cky/"&imagename
	   end if
	End If	
	
	infoBoMe=	Trim(Upload.Form("infoBoMe"))
	infoBoMe=	Replace(infoBoMe,"<br>",chr(13) & chr(10))			
	infoAnhEm=	Trim(Upload.Form("infoAnhEm"))
	infoAnhEm=	Replace(infoAnhEm,"<br>",chr(13) & chr(10))
	infoVoChongCon=Trim(Upload.Form("infoVoChongCon"))
	infoVoChongCon=	Replace(infoVoChongCon,"<br>",chr(13) & chr(10))
	Hoatdongbanthan=Trim(Upload.Form("Hoatdongbanthan"))
	Hoatdongbanthan=	Replace(Hoatdongbanthan,"<br>",chr(13) & chr(10))

	TK_NganHang=Upload.Form("TK_NganHang")
	BankID	=	Upload.Form("BankID")
	

	set eChuKyfile = Upload.Files("eChuKyFile")
	If eChuKyfile Is Nothing Then
		eChuKy=""
	else
	   Filetype = Right(eChuKyfile.Filename,len(eChuKyfile.Filename)-Instr(eChuKyfile.Filename,"."))
	   if   Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" then
			sError=sError & "&nbsp;-&nbsp; Không phải dạng chữ ký<br>"
	   else
			imagename="CK"&CMT& "." & Filetype			
			eChuKyfile.SaveAs server.MapPath("/administrator/images/Cky/")& "\" & imagename
			eChuKy	=	"/administrator/images/Cky/"&imagename
	   end if
	End If		

	RemoveImageimgNV		=GetNumeric(Upload.Form("RemoveImageimgNV"),0)
	if RemoveImageimgNV = 1 then
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(server.MapPath(Upload.Form("imgNV"))) then
		  fs.DeleteFile(server.MapPath(Upload.Form("imgNV")))
		end if
		set fs=nothing
	end if
	
	RemoveImageeChuKy		=GetNumeric(Upload.Form("RemoveImageeChuKy"),0)
	if RemoveImageeChuKy = 1 then
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(server.MapPath(Upload.Form("imgESigna"))) then
		  fs.DeleteFile(server.MapPath(Upload.Form("imgESigna")))
		end if
		set fs=nothing
	end if	

	if iStatus="edit" then 
		StaffContractID	=	GetNumeric(Upload.Form("StaffContractID"),0)	
		sql = "UPDATE Nhanvien SET "
		sql = sql & "Ho_Ten = N'"&Ho_Ten&"'"
		sql = sql & ",NgaySinh= '"&NgaySinh&"'"
		sql = sql & ",CMT= '"&CMT&"'"
        sql = sql & ",salecode= '"&salecode&"'"
		sql = sql & ",NgayCap= '"&NgayCap&"'"
		sql = sql & ",Noicap= N'"&Noicap&"'"
		sql = sql & ",Hocvan= N'"&Hocvan&"'"
		sql = sql & ",DanToc= N'" & DanToc&"'"
		sql = sql & ",Tel= N'"&Tel&"'"
		sql = sql & ",Mobile= N'"&Mobile&"'"
		sql = sql & ",Email= '"&Email&"'"
		sql = sql & ",Cu_tru= N'"&Cu_tru&"'"
		sql = sql & ",Diachi= N'"&Diachi&"'"
		if imgNV<>""  then	
			sql = sql & ",imgNV= '"&imgNV&"'"
		elseif RemoveImageimgNV = 1 then
			sql = sql & ",imgNV= ''"
		end if
		sql = sql & ",infoBoMe= N'"&infoBoMe&"'"
		sql = sql & ",infoAnhEm= N'"&infoAnhEm&"'"
		sql = sql & ",infoVoChongCon= N'" & infoVoChongCon&"'"
		sql = sql & ",Hoatdongbanthan= N'"&Hoatdongbanthan&"'"
		sql = sql & ",TK_NganHang= '"&TK_NganHang&"'"
		sql = sql & ",BankID= '"&BankID&"'"
		if eChuKy<>""  then	
			sql = sql & ",eChuKy= '"&eChuKy&"'"
		elseif RemoveImageeChuKy = 1 then
			sql = sql & ",eChuKy= ''"
		end if
		sql = sql & ",password_partner= '"&password_partner&"'"	
		sql = sql & " where NhanVienID = "&StaffContractID
	else
		sql = "INSERT INTO Nhanvien(Ho_Ten, NgaySinh, CMT, NgayCap, Noicap, Hocvan, DanToc, Tel, Mobile, Email, Cu_tru, Diachi, imgNV, infoBoMe, infoAnhEm, infoVoChongCon, Hoatdongbanthan, TK_NganHang, BankID, eChuKy, salecode,password_partner)"
		sql = sql & " VALUES ( "
		sql = sql & "N'"&Ho_Ten&"'"
		sql = sql & ", '"&NgaySinh&"'"
		sql = sql & ", '"&CMT&"'"
		sql = sql & ", '"&NgayCap&"'"
		sql = sql & ", N'"&Noicap&"'"
		sql = sql & ", N'"&Hocvan&"'"
		sql = sql & ", N'" & DanToc&"'"
		sql = sql & ", N'"&Tel&"'"
		sql = sql & ", N'"&Mobile&"'"
		sql = sql & ", '"&Email&"'"
		sql = sql & ", N'"&Cu_tru&"'"
		sql = sql & ", N'"&Diachi&"'"
		sql = sql & ", '"&imgNV&"'"
		sql = sql & ", N'"&infoBoMe&"'"
		sql = sql & ", N'"&infoAnhEm&"'"
		sql = sql & ", N'" & infoVoChongCon&"'"
		sql = sql & ", N'"&Hoatdongbanthan&"'"
		sql = sql & ", '"&TK_NganHang&"'"
		sql = sql & ", '"&BankID&"'"
		sql = sql & ", '"&eChuKy&"'"
        sql = sql & ", '"&salecode&"'"
		sql = sql & ", '"&password_partner&"')"
	end if
elseif Catup = "contract" then
	NgayHD=	Trim(Upload.Form("NgayHD"))
	ThangHD=	trim(UpLoad.Form("ThangHD"))
	NamHD	=	Trim(UpLoad.Form("NamHD"))
	NgayHD	=	ThangHD&"/"&NgayHD&"/"&NamHD
	if isdate(NgayHD) = true then
		NgayHD	=FormatDatetime(NgayHD)
	else
		NgayHD = Null
	end if	
	
	NgayEnd	=	Trim(Upload.Form("NgayEnd"))
	ThangEnd=	trim(UpLoad.Form("ThangEnd"))
	NamEnd	=	Trim(UpLoad.Form("NamEnd"))
	NgayEnd	=	ThangEnd&"/"&NgayEnd&"/"&NamEnd
	if isdate(NgayEnd) = true then
		EndDate	=FormatDatetime(NgayEnd)
	else
		EndDate = Null
	end if		
	
	ChucdanhID	=	GetNumeric(Upload.Form("ChucdanhID"),0)
	PhongID		=	GetNumeric(Upload.Form("PhongID"),0)
	isMember	=	GetNumeric(Upload.Form("isMember"),0)
	luong		=	GetNumeric(Chuan_money(Upload.Form("luong")),0)
    moneyoff		=	GetNumeric(Chuan_money(Upload.Form("moneyoff")),0)

	Heso		=	GetNumeric(Upload.Form("Heso"),0)
	trachnhiem	=	GetNumeric(Chuan_money(Upload.Form("trachnhiem")),0)
	Phucap		=	GetNumeric(Chuan_money(Upload.Form("Phucap")),0)
	
	luong_BH	=	GetNumeric(Chuan_money(Upload.Form("luong_BH")),0)
	Congviec	=	Trim(Upload.Form("Congviec"))
	Congviec	=	Replace(Congviec,"<br>",chr(13) & chr(10))
	
	Dieukhoan	=	Trim(Upload.Form("Dieukhoan"))
	Dieukhoan	=	Replace(Dieukhoan,"<br>",chr(13) & chr(10))
	
	StaffSigna	=	GetNumeric(Upload.Form("StaffSigna"),0)
	ThuTruongID	=	GetNumeric(Upload.Form("ThuTruongID"),0)
	DongY		=	GetNumeric(Upload.Form("CDongY"),0)
	
	RemoveImageeChuKy		=GetNumeric(Upload.Form("RemoveImageeChuKy"),0)
	if RemoveImageeChuKy = 1 then
		StaffSigna	 = 0
	end if
	RemoveImageeChuKy		=GetNumeric(Upload.Form("RemoveImageeChuKy0"),0)
	if RemoveImageeChuKy = 1 then
		Dongy	 = 0
	end if	
	
	if iStatus="edit" then 
		ContractID	=	GetNumeric(Upload.Form("ContractID"),0)	
		sql = "UPDATE StaffContract SET "
		sql = sql & "NgayHD = '"&NgayHD&"'"
		sql = sql & ",EndDate= '"&EndDate&"'"
		sql = sql & ",ChucdanhID= '"&ChucdanhID&"'"
		sql = sql & ",PhongID= '"&PhongID&"'"
		sql = sql & ",isMember= '"&isMember&"'"
        sql = sql & ",moneyoff= '"&moneyoff&"'"
		sql = sql & ",luong= '"&luong&"'"
		sql = sql & ",heso= '" & heso&"'"
		sql = sql & ",trachnhiem= '"&trachnhiem&"'"
		sql = sql & ",luong_BH= '"&luong_BH&"'"
		sql = sql & ",Phucap= '"&Phucap&"'"
		sql = sql & ",ThuTruongID= '"&ThuTruongID&"'"
		sql = sql & ",Dongy= '"&Dongy&"'"
		sql = sql & ",Kynhan='"&StaffSigna&"'"
		sql = sql & ",Congviec= N'"&Congviec&"'"
		sql = sql & ",Dieukhoan= N'"&Dieukhoan&"'"
		sql = sql & " where ID = "&ContractID
	else
		NhanVienID	=	GetNumeric(Upload.Form("NhanVienID"),0)
		sql = "INSERT INTO StaffContract( NhanVienID, NgayHD, EndDate, ChucdanhID, PhongID, isMember, moneyoff,luong, heso, trachnhiem, luong_BH, Phucap, ThuTruongID, Dongy, Kynhan, Congviec,Dieukhoan)"
		sql = sql & " VALUES ( "
		sql = sql & " '"&NhanVienID&"'"
		sql = sql & ", '"&NgayHD&"'"
		sql = sql & ", '"&EndDate&"'"
		sql = sql & ", '"&ChucdanhID&"'"
		sql = sql & ", '"&PhongID&"'"
		sql = sql & ", '"&isMember&"'"
        sql = sql & ", '"&moneyoff&"'"
		sql = sql & ", '" & luong&"'"
		sql = sql & ", '"& Heso &"'"
		sql = sql & ", '"&trachnhiem&"'"
		sql = sql & ", '"&luong_BH&"'"
		sql = sql & ", '"&Phucap&"'"
		sql = sql & ", '"&ThuTruongID&"'"
		sql = sql & ", '"&Dongy&"'"
		sql = sql & ", '"&StaffSigna&"'"
		sql = sql & ", N'"&Congviec&"'"
		sql = sql & ", N'"&Dieukhoan&"')"
	end if
end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,3
%>
<script language="javascript">
	alert('Da cập nhật xong thua Ong chu!');
	history.back();
	window.history.reload();	
</script>

