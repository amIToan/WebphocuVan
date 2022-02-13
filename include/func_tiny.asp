
<%
sub Day_morth_year(ngay,thang,nam)
%>

<select name="select_day" id="select_day">
			
			<% 
			i =1
			for i=1 to 31
			if i=ngay then
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option>"& i &"</option>")
			next
			%>
</select> 
          	<select name="select_month" id="select_month">
            
			<% 
			i =1
			for i=1 to 12
			if i=thang then
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option>"& i &"</option>")
			next
			%>
          </select>
           <select name="select_year" id="select_year">
		   
		   <%
		   	i=1
			for i=2004 to 2010 
			if i=nam then
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if			
			Response.Write("<option>"& i &"</option>")
			 
            next
         	%>
          </select>		
<%

end sub

%>

  <%
sub ex_Day_morth_year(ngay,thang,nam,dayValue,morValue,yearValue)
%>
<select name="<%=dayValue%>" id="<%=dayValue%>">
			
			<% 
			i =1
			for i=1 to 31
			if i=ngay then
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option>"& i &"</option>")
			next
			%>
</select> 
          	<select name="<%=morValue%>" id="<%=morValue%>">
           	<% 
			i =1
			for i=1 to 12
			if i=thang then
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if
			if i<10 then
			i="0"&i
			end if
			Response.Write("<option>"& i &"</option>")
			next
			%>
          </select>
           <select name="<%=yearValue%>" id="<%=yearValue%>">
		   <%
		   	i=1
			for i=1900 to 2010 
			if i=nam then
			Response.Write("<option selected='selected'>"& i &"</option>")
			i=i+1
			end if			
			Response.Write("<option>"& i &"</option>")
	        next
         	%>
          </select>		
<%
end sub
%>

<%function return_day()

ngay=Request.Form("select_day")
if ngay<10 then
		ngay="0" & ngay
end if
thang=Request.Form("select_morth")
if thang<10 then
		thang="0" & thang
end if
nam=Request.Form("select_year")
if nam<10 then
		nam="0" & nam
end if

return_day=ngay&"/"&thang&"/"&nam

end  function

function fun_convert_day(date_time)
      		k=len(date_time)
			i = InStr(date_time, "/")
       		thang = Mid(date_time,1, i - 1)
			if Cint(thang)<10 then
			thang="0"& thang
			end if			
      	 	date_time = Mid(date_time,i+1, k)
        	i = InStr(date_time, "/") 
			ngay = Mid(date_time,1, i - 1)
			if Cint(ngay)<10 then
			ngay="0"& ngay
			end if	
       		date_time = Mid(date_time,i+1,k)
			nam = date_time
			if Cint(nam)<10 then
			nam="0"& nam
			end if	
			fun_convert_day=ngay+"/"+thang+"/"+nam			
 End function
%>

<%
sub fundate_one_three(date_time)
			k=len(date_time)
			i = InStr(date_time, "/")
       		thang = Mid(date_time,1, i - 1)			
      	 	date_time = Mid(date_time,i+1, k)
        	i = InStr(date_time, "/") 
			ngay = Mid(date_time,1, i - 1)
       		date_time = Mid(date_time,i+1,k)
			nam = date_time
			ngay=Cint(ngay)
			thang=Cint(thang)
			nam=Cint(nam)
end sub
%>

<%
function tienchu(sotien)
	if sotien<0 then
		chu=chu &"âm "
	end if 
	sotien=ABS(sotien)
	IF sotien>999999999999 then 
		msgbox"Số quá lớn"
		Exit function 
	end if 
	if sotien = 0 then
		chu = "Không đồng"
	else
		tmpchuoi =Trim(cstr(sotien))
		iLen = len(tmpchuoi)
		for iTemp=1 to iLen
			inumber = Cint(Mid(tmpchuoi,iTemp,1))
			iPosition = iLen-iTemp	
			chu = chu & doisochu(inumber,iPosition)
			chu = Replace(chu,"lẻ đồng", "đồng")
			chu = Replace(chu,"lẻ triệu", "triệu")
			chu = Replace(chu,"lẻ nghìn", "nghìn")
			chu = Replace(chu,"lẻ tỷ", "tỷ")
		next
	end if
	tienchu=chu
End Function


Function doisochu(inumber,i)
	chuoi=""
	donvi=""
	select case i
		case 9
			donvi=" tỷ "
		case 6
			donvi=" triệu "
		case 3
			donvi=" nghìn "
		case 2,5,8,11
			donvi=" trăm "
		case 0
			donvi=" đồng "
		case 1,4,7,10
			donvi=" mươi "
			if inumber =0 then
				donvi= " lẻ "
			end if
			if inumber = 1 then
				donvi=" mười "
			end if
	end select
	select case inumber
		case 0
			chuoi=" không "
			if donvi=" lẻ " then
				chuoi=""
			end if
		case 1
			chuoi=" một "
			if donvi=" mười " then
				chuoi=""
			end if
		case 2
			chuoi=" hai"
		case 3
			chuoi=" ba"
		case 4
			chuoi=" bốn"
		case 5
			chuoi=" năm"
		case 6
			chuoi=" sáu"
		case 7
			chuoi=" bảy"
		case 8
			chuoi=" tám"
		case 9
			chuoi=" chín"
	end select
	str=chuoi&donvi
	if inumber = 0  then ' cắt bỏ ký tự "không" và đơn vị
		select case i
			case 0
				str	=	" đồng"
			case 3
				str = 	" nghìn"
			case 6
				str = 	" triệu"
			case 9
				str	=	" tỷ"
		end select
	end if
	doisochu=str
End Function
%>

<%
function chuan_money(txt_money)
	dim str_temp
	k_lengh=len(txt_money)
	if left(txt_money,1) = "-" then
		str_temp = "-"
	end if
	for bieni=1 to k_lengh
		str=mid(txt_money,bieni,1)
		if isNumeric(str) then
			str_temp=str_temp+str
		end if
	next
	chuan_money=str_temp
end function
function GetidcodeFromNewsID(NewsID)
	sql = "SELECT idcode FROM News WHERE NewsID = '"& NewsID &"'"
	set rsIDSP = Server.CreateObject("ADODB.recordset")
	rsIDSP.open sql,con,1
	if not rsIDSP.EOF then
		GetidcodeFromNewsID = rsIDSP("idcode")
	else
		GetidcodeFromNewsID = ""
	end if	
	set rsIDSP =nothing
end function

function GetCuocBuuDien(Trongluong)
	giabd=0
	if Trongluong <= 250 then
		giabd =  7500	
	elseif Trongluong > 250 and Trongluong <= 500 then
		giabd =  9500
	elseif Trongluong > 500 and Trongluong <= 1000 then
		giabd =  15000
	elseif Trongluong > 1000 and Trongluong <= 1500 then
		giabd =  18500
	elseif Trongluong > 1500 and Trongluong <= 2000 then	
		giabd =  21500
	else
		TlConlai = Trongluong - 2000
		sdu = TlConlai mod 500
		hso = (TlConlai - sdu)/500
		if sdu > 0 then
			hso = hso + 1
		end if
		giabd = hso*2000 + 21500
	end if
	GetCuocBuuDien = giabd
end function

function GetCuocBuuDienNhanh(Trongluong,KhoangCach)
	giabd=0
	if Trongluong <= 50 then
		if KhoangCach = 0 then
			giabd =  10973
		else
			giabd =  11550
		end if
	elseif Trongluong > 50 and Trongluong <= 100 then
		if KhoangCach = 0 then
			giabd =  14438
		else
			giabd =  16170
		end if
	elseif Trongluong > 100 and Trongluong <= 250 then
		if KhoangCach = 0 then
			giabd =  19058
		else
			giabd =  23100
		end if	
	elseif Trongluong > 250 and Trongluong <= 500 then
		if KhoangCach = 0 then
			giabd =  26565
		else
			giabd =  30608
		end if	
	elseif Trongluong > 500 and Trongluong <= 1000 then
		if KhoangCach = 0 then
			giabd =  37508
		else
			giabd =  44468
		end if
	elseif Trongluong > 1000 and Trongluong <= 1500 then
		if KhoangCach = 0 then
			giabd =  46200
		else
			giabd =  57173
		end if
	elseif Trongluong > 1500 and Trongluong <= 2000 then	
		if KhoangCach = 0 then
			giabd =  54863
		else
			giabd =  68732
		end if
	else
		TlConlai = Trongluong - 2000
		sdu = TlConlai mod 500
		hso = (TlConlai - sdu)/500
		if sdu > 0 then
			hso = hso + 1
		end if
		if KhoangCach = 0 then
			giabd = hso*5000 + 54863
		else
			giabd =  hso*6000 + 68732
		end if		
	end if
	giabd	=	giabd + 0.1*giabd
	GetCuocBuuDienNhanh = giabd
end function

function CheckSanPhamNhap(NewsID)
	sql = "select NewsID From SanPhamNhap where NewsID = '"& NewsID &"'"
	Set rs1=server.CreateObject("ADODB.Recordset")
	rs1.open sql,con,1
	if not rs1.eof then
		CheckSanPhamNhap = true
	else
		CheckSanPhamNhap = false
	end if
	set rs1 = nothing
end function

'tuannv 16/7/2008.
Function GetItemParameters(iItemID)
Dim aParameters 
	sqlNews="SELECT * from V_News where NewsId=" & iItemID
	Set rsTemp=Server.CreateObject("ADODB.Recordset")
	rsTemp.open sqlNews,con,3
	if rsTemp.eof then
		rsTemp.close
		set rsTemp=nothing
		aParameters = Array("", "", "")
		exit Function
	else
		idcode   =   rsTemp("idcode")
		Title		=	rsTemp("Title")
		if Title <> "" then
			Title=Trim(replace(Title,"'","''"))
			Title=Replace(Title,"""","&quot;")
		end if
		PictureId	=	rsTemp("PictureId")
		nxb			=	rsTemp("nxb")
		if Trim(rsTemp("tacgia"))<>"" then
			tacgia=Trim(rsTemp("tacgia"))
			tacgia=Replace(tacgia,"'","''")
			tacgia=Replace(tacgia,chr(13) & chr(10),"<br>")
		else
			tacgia=""
		end if
		namxuatban	=	rsTemp("namxuatban")
		if  len(namxuatban) >4 then
			namxuatban = left(namxuatban,4)
		else
			namxuatban=0
		end if
		Giabia		=	rsTemp("Giabia")
		Gia			=	rsTemp("Gia")
		trongluong	=	rsTemp("trongluong")
		'						0		1    2    3          4      5      6     7         8           9
		aParameters = Array(iItemID, Title, Gia,PictureId, Giabia,tacgia, nxb,namxuatban,trongluong,idcode)
		rsTemp.close
		set rsTemp=nothing
	end if
GetItemParameters = aParameters
End Function

Function GetMaxId(TableName, FieldNameId, sCondition)
	Dim Max,rsMaxId
	set rsMaxId=server.CreateObject("ADODB.Recordset")
	sql="select Max(" & FieldNameId & ") as MaxId from " & TableName
	if sCondition<>"" then
		sql=sql & " where sCondition"
	end if
	rsMaxId.Open sql, con, 1
	if IsNull(rsMaxId("MaxId")) then
		Max=1
	else
		Max=CLng(rsMaxId("MaxId")) +1
	end if
	rsMaxId.close
	set rsMaxId=nothing
	GetMaxId=Max
End Function

function setStyleBorder(isLeft,isRight,isTop,isBottom)
	strTemp = ""	
	if isLeft = 1 then
		strTemp = strTemp +  "border-left:#99CCFF solid 1px;"
	end if
	if isRight = 1 then
		strTemp = strTemp +  "border-Right:#99CCFF solid 1px;"
	end if
	if isTop = 1 then
		strTemp = strTemp +  "border-top:#99CCFF solid 1px;"
	end if
	if isBottom = 1 then
		strTemp = strTemp +  "border-bottom:#99CCFF solid 1px;"
	end if
	if isLeft = 1 and isRight = 1 and isTop = 1 and isBottom = 1 then
		strTemp = "border:#99CCFF solid 1px;"
	end if
	setStyleBorder= strTemp
end function



function GetTenTinh(ID)
	sql = "Select TenTinh FROM Tinh where TinhID="&ID
	Set rsTinh=Server.CreateObject("ADODB.Recordset")
	rsTinh.open sql,con,3
	if not rsTinh.eof then
		GetTenTinh = rsTinh("TenTinh")
	else
		GetTenTinh = ""
	end if
	set rsTinh = nothing
end function

function GetTenHuyen(ID)
	sql = "Select TenHuyen FROM Huyen where HuyenID="&ID
	Set rsHuyen=Server.CreateObject("ADODB.Recordset")
	rsHuyen.open sql,con,3
	if not rsHuyen.eof then
		GetTenHuyen = rsHuyen("TenHuyen")
	else
		GetTenhuyen = ""
	end if
	set	rsHuyen = nothing
end function

%>
<%
function GetThuTrongTuan(Thu)
	strThu = ""
	select  case Thu
	case 1
		strThu = "Chủ nhận"
	case 2
		strThu = "Thứ 2"
	case 3
		strThu = "Thứ 3"
	case 4
		strThu = "Thứ 4"
	case 5
		strThu = "Thứ 5"
	case 6
		strThu = "Thứ 6"
	case 7
		strThu = "Thứ 7"
	end select
	GetThuTrongTuan = strThu
end function 
%> 

<%
function MaHoaKyTu(KyTu)
ktTemp	=	KyTu
MaHoaKyTu	=	ktTemp
end function

function GiaiMaKyTu(KyTuMaHoa)
ktTemp	=	KyTuMaHoa
GiaiMaKyTu	=	ktTemp	
end function

function GetTextBox(txtName)
	txtName	=	Trim(Request.Form(txtName))
	txtName=Replace(txtName,"'","''")
	txtName=Replace(txtName,chr(13) & chr(10),"<br>")
	GetTextBox	=	txtName
end function

function AddOnline(iVitri)
	iVt = iVitri
	strIP	=	Request.ServerVariables("REMOTE_ADDR")
	iTime	=	now()
	iSubTime	=	DateAdd("n",-30,iTime)
	sql="select Count(ID) as iCount from MOnline where ip ='"& strIP &"' and TimeDays > '"& iSubTime &"'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	on error Resume next
	rs.open sql,Con,3
	iCount =	rs("iCount")
	set rs = nothing
	if iCount = 0 then
		sql="insert into MOnline(ip,TimeDays,iVitri) values('"& strIP &"','"& iTime &"','"& iVt &"')"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,Con,3
		set rs = nothing
	end if
end function
%>

<%Function GetFullDate(theDate)
	ngay = CInt(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = CInt(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = CInt(year(theDate))
	GetFullDate=ngay & "/" & thang & "/" & nam
End Function%>

<%
function GetTacGia(NewsID)
	sql="SELECT  top 1 tacgia FROM V_News where NewsID='"& NewsID &"'" 
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	rsTemp.open sql,Con,3
	strTacgia = ""
	if not rsTemp.eof then
		strTacgia	=	rsTemp("tacgia")
	end if
	strTacgia	=	replace(strTacgia,"'s","")
	'cắt chỉ lấy phần của tác giả
	iP	=	Instr(strTacgia,"- Dịch giả:")
	if iP > 5 then
	strTacgia	=	Mid(strTacgia,1,iP-1)
	end if
	GetTacGia	=	strTacgia
end function 
%>
<%
function GetTieuDeSachTB(ID)
	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	strTieuDe	=	""
	sql =	"Select * from SachTieuBieu where ID="&ID
	rsTemp.open sql,con,1
	if  not rsTemp.EOF then
		strTieuDe	=	Trim(rsTemp("TieuDe"))
	end iF
	set rsTemp = nothing	
	GetTieuDeSachTB	=strTieuDe	
end function

function isKichHoatSachTB(ID)
	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	iKichHoat	= false
	sql =	"Select * from SachTieuBieu where ID="&ID
	rsTemp.open sql,con,1
	iF not rsTemp.EOF then
		if  rsTemp("KichHoat") <> 0 then
			iKichHoat	=	true
		end if
	end if				
	set rsTemp = nothing	
	isKichHoatSachTB	=iKichHoat
end function 
function GetIDKichHoatSachChuDe()
	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	ID	= 0
	sql =	"Select Top 1 ID from SachTieuBieu where KichHoat=1"
	rsTemp.open sql,con,1
	iF not rsTemp.EOF then
		ID	=	GetNumeric(rsTemp("ID"),0)
	end if				
	set rsTemp = nothing	
	GetIDKichHoatSachChuDe	=ID	
end function
%>
<%
function VanHoaText(text)
	bTrueFalse	=	true
	arKoVH	=	Array("mẹ mày","me may","chung may","chúng mày","địt","dit","tien su","tiên sư","chó chết","cho chet","đểu cán","deu can","bon tao","bọn tao","tui bay","tụi bây","tui may","tụi mày")
	for k = 0 to Ubound(arKoVH)
		if InStr(Ucase(text),Ucase(arKoVH(k))) >= 1 then
			bTrueFalse	=	 false
		end if
	next
	VanHoaText	=	bTrueFalse		
end function
%>
<%
    function fOnlyCallName(sName)
        xungho  =   fXungHo(sName)
        sName   =   Trim(sName)
        Pos     = InStrRev (sName," ")
        sName   = Mid(sName,Pos+1)
        
        fOnlyCallName   = xungho & " " & sName
    end function

     %>

<%
function fXungHo(Ten)
	if Ten <> "" then
	arNam	=	Array("Ông","Mr","Văn","Nam","Hưng","Hùng","Quang","Tuấn","Hiếu","Thành","Quân","Dũng","Duy","Tùng","Trí","Sơn","Trưng","Thăng","Thắng","Việt","Trường","Tiến","Toàn","Đức","Cường","Hiệp","Long","Thọ","Khoa","Trọng","Công","Cương","Phong","Kiên","Hưng","Huy","Bộ","Đông","Thọ","Lâm","Mạnh","Xây","Đạt","Hữu","Thịnh","Sĩ","Đình","Tiệp","Tuân","Bá","Tân","Tấn","Đại","Vĩnh","Tính","Dầu","Tý","Tam","Lục","Trung")
	
	arNu	=	Array("Miss","Ms","Bà","Chị","Thị","Thi","Gái","Nữ","Hằng","Hương","Chi","Trang","Thúy","Lan","Ngân","Huyền","Liên","Huệ","Đào","Liễu","Nhung","Trúc","Mai","Nga","Quyên","Lệ","Yến","Nguyệt","Bích","Hoài","Hảo","Thảo","Thơm","Loan","Uyên","Lành","Tươi","Dung","Vy","Hoa","Diệp","Điệp","Ly","Diệu","Thắm","Nhi","Duyên","Thêu","Thủy","Tiên","Kim","Khánh","Phượng","Gấm")
	
	arNamNu	=	Array("Anh","Ánh","Chung","Ngọc","Hải","Giang","Châu","Phương","Thu","Hà","Minh","Yên"," Thương","Vui","Thụy","Sự","Thu","Bình","Hạnh","Xuân","Thuận","Bạn","Thùy","Thanh","Nguyên","Dương","Tú","Châu","Sáu","Bảy","Hai","Tư","Quỳnh")
	
	isNam	=	0
	isNu	=	0
	isNamNu	=	0

	TenTemp=	Ucase(Ten)
	arTen	=	split(TenTemp," ")
	lTen	=	UBound(arTen)
	for k = 0 to lTen
		for t=0 to Ubound(arNam) 
			if StrComp(Ucase(arNam(t)),arTen(k)) =0 then
				isNam	=	isNam	+ 	1
			end if
		next
		
		for t = 0 to UBound(arNu)
			if StrComp(Ucase(arNu(t)),arTen(k)) = 0 then
				isNu	=	isNu	+ 	1
			end if
		next
		for t = 0 to UBound(arNamNu)
			if StrComp(Ucase(arNamNu(t)),arTen(k))=0 then
				isNamNu	=	isNamNu	+ 	1
			end if
		next
	next
	'Response.Write(isNam&" "&isNu&" "&isNamNu)
	if isNam > isNu then
		xungho	=	"anh"
	elseif isNu > isNam then
		xungho	=	"chị"
	elseif isNam = isNu and isNam <> 0 then
		xungho	=	"bạn"
	elseif isNamNu >= 3 then
		xungho	=	"chị"
	else
		xungho	=	"bạn"		
	end if
		
	fXungHo	=	xungho		
	else
		fXungHo=""
	end if
end function
%>

<%
sub SelectProvider(NameProvider,ProvideID)
	strPviID="Select ProviderID,ProviderName From Provider order by ProviderName"
	dim rsPrVID
	set rsPrVID=Server.CreateObject("ADODB.Recordset")
	rsPrVID.open strPviID,Con,3 	
	%>
	<select name="<%=NameProvider%>" size="1" id="<%=NameProvider%>">
	<option value="0">Mời chọn</option>
	<% do while not rsPrVID.eof%>
	<option value="<%=rsPrVID("ProviderID")%>"<%if ProvideID=rsPrVID("ProviderID") then  Response.Write("selected=""selected""") end if %>>
	<%=rsPrVID("ProviderName")%>
	</option>
	<% 
	rsPrVID.movenext
	loop
	set strPviID=nothing
	%>
	</select>
<%	
end sub
%>

<%
function GetNumInGoodsStore(NewsID)
	sql = " SELECT SUM(Number) AS tNumNhap FROM Product WHERE NewsID="&NewsID
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetNumInGoodsStore = GetNumeric(rsTemp("tNumNhap"),0)
	else
		GetNumInGoodsStore = 0
	end if
	set rsTemp = nothing
end function

function GetNumOutGoodsStore(NewsID)
	sql = " SELECT  SUM(XuatKho.SoLuong) AS iCount FROM XuatKho INNER JOIN Product ON XuatKho.ProductID = Product.ProductID WHERE Product.NewsID = "&NewsID
	set rsTemp =  Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetNumOutGoodsStore = GetNumeric(rsTemp("iCount"),0)
	else
		GetNumOutGoodsStore=0
	end if
	set rsTemp = nothing
end function

function GetNumInventoryGoods(NewsID)
	iTonXbook	= 	0
	iTonXbook = GetNumInGoodsStore(NewsID)-GetNumOutGoodsStore(NewsID)-GetNumReturnGoods(NewsID)
	'Response.Write("NHập:"&GetNumInGoodsStore(NewsID)&"; xuất:"&GetNumOutGoodsStore(NewsID)&"; trả:"&GetNumReturnGoods(NewsID))
	GetNumInventoryGoods	=	iTonXbook
end function

function GetNumReturnGoods(NewsID)
	sql	=	"SELECT     SUM(TraSach.SLTraNCC) AS iSum FROM Product INNER JOIN TraSach ON Product.ProductID = TraSach.ProductID WHERE (Product.NewsID = '"& NewsID &"')"
	Set rsTemp	=	Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetNumReturnGoods = GetNumeric(rsTemp("iSum"),0)
	else
		GetNumReturnGoods=0
	end if
	set rsTemp = nothing
end function


function GetPoint(SanPhamUser_ID)
	iPoint=0
	sql = "SELECT  GetPoints FROM  SanPham_pay  WHERE SanPhamUser_ID="&SanPhamUser_ID
	Set rsPoint=Server.CreateObject("ADODB.Recordset")
	rsPoint.open sql,con,1
	if not rsPoint.eof then
		iPoint	=	rsPoint("GetPoints")
	end if
	GetPoint	=	iPoint	
end function
%>


<%
sub OtherIdea(NewsID)
	sqlNews="SELECT * FROM Y_KIEN WHERE NewsId='"& NewsId &"' and show = '1'"
	Set rs_idea=Server.CreateObject("ADODB.Recordset")
	rs_idea.open sqlNews,con,1
	if not rs_idea.eof then
%>	
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2" background="../../images/TabChinh.gif"  style="background-position:bottom left; background-repeat:no-repeat;" class="CTieuDeNho"><img src="../../images/bullet277.gif" width="32" height="32"  align="absmiddle" />&nbsp;&nbsp;&nbsp;Ý Kiến Khách Hàng</td>
  </tr>
<%
Do while not rs_idea.EOF	
	hovaten=rs_idea("hovaten")
	tieude=rs_idea("tieude")
	noidung=rs_idea("noidung")
%>  
  <tr>
    <td width="148" style="<%=setStyleBorder(1,1,0,1)%>"><img src="../../images/icons/icon_news.gif" width="32" height="32" align="absmiddle" />&nbsp;<%=hovaten%></td>
    <td width="422" align="justify" style="<%=setStyleBorder(0,1,0,1)%>">
	<font class="CTieuDeNhoNho">&nbsp;<%=tieude%></font><br />
	<font class="CTxtContent">&nbsp;<%=noidung%></font>
		  	
	</td>
  </tr>
<%
	rs_idea.MoveNext
Loop
rs_idea.Close
set rs_idea=nothing
%>  
  
</table>	
	
<%	
	end if
end sub	
%>


<%
  Function SortDictionary(objDict,intSort,isnumber)
    Dim strDict()
    Dim objKey
    Dim strKey,strItem
    Dim X,Y,Z
    Z = objDict.Count
    If Z > 1 Then
      ReDim strDict(Z,2)
      X = 0
      For Each objKey In objDict
          strDict(X,1)  = objKey
          strDict(X,2) = objDict(objKey)
          X = X + 1
      Next
      For X = 0 to (Z - 2)
        For Y = X to (Z - 1)
          If ((StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 and isnumber = 0) or (strDict(X,intSort) < strDict(Y,intSort) and isnumber = 1)) then
              strKey  = strDict(X,1)
              strItem = strDict(X,2)
              strDict(X,1)  = strDict(Y,1)
              strDict(X,2) = strDict(Y,2)
              strDict(Y,1)  = strKey
              strDict(Y,2) = strItem
          End If
        Next
      Next
      objDict.RemoveAll
      For X = 0 to (Z - 1)
        objDict.Add strDict(X,1), strDict(X,2)
      Next
    End If
  End Function
%>

<%
function GetCategoryLoai(CatID)
	sql = "Select CategoryLoai from NewsCategory where CategoryID ="&CatID
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if not rs.eof then	
		GetCategoryLoai	=	rs("CategoryLoai")
	end if
end function
%>

<%
function GetTitleNews(NewsID)
	sql = "Select Title from News where NewsID ="&NewsID
	Set rs_get_title=Server.CreateObject("ADODB.Recordset")
	rs_get_title.open sql,con,1
	if not rs_get_title.eof then	
		GetTitleNews	=	rs_get_title("Title")
    else
        GetTitleNews=""
	end if
    set rs_get_title=nothing
    
end function
%>


<%Function GetNameOfCategoryLoai(CategoryLoai)
	Select case CategoryLoai
		case 0
			GetNameOfCategoryLoai="Giới thiệu"
		case 1
			GetNameOfCategoryLoai="Tin tức"
		case 2
			GetNameOfCategoryLoai="FAQ"
		case 3
			GetNameOfCategoryLoai="Sản phẩm"
		case 4
			GetNameOfCategoryLoai="Liên hệ"
		case 5
			GetNameOfCategoryLoai="Download"	
		case 6
			GetNameOfCategoryLoai="Tìm kiếm"	
		case 7
			GetNameOfCategoryLoai="Văn bản"	
		case 8
			GetNameOfCategoryLoai="Thư giãn"	
		case 9
			GetNameOfCategoryLoai="Hướng dẫn"
        case 10
            GetNameOfCategoryLoai="Tư vấn"
	End Select



End Function

Function Is_Mobile()
  Set Regex = New RegExp
  With Regex
    .Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad)"
    .IgnoreCase = True
    .Global = True
  End With
  Match = Regex.test(Request.ServerVariables("HTTP_USER_AGENT"))
  If Match then
    Is_Mobile = True
  Else
    Is_Mobile = False
  End If
End Function

%>


<%Sub ListLoaiOfCategory(CategoryLoai)%>
<%
sql="Select * from CategoryLoai"
Set rstemp  = server.CreateObject("ADODB.Recordset")
rsTemp.open sql,con,1
 %>
	<Select name="CategoryLoai" id="CategoryLoai">
<%
do while not rsTemp.eof
 %>
		<option value="<%=rstemp("CategoryLoai")%>"<%if Cint(CategoryLoai)=Cint(rstemp("CategoryLoai"))then%> selected <%End if%>><%=rstemp("TxtCategory") %></option>
<%
    rsTemp.movenext
loop %>
	</Select>
<%End Sub%>

<%Function GetNameOfCategoryLoai(CategoryLoai)
	sql="Select * from CategoryLoai where CategoryLoai = '"&CategoryLoai&"'"
	Set rstemp12  = server.CreateObject("ADODB.Recordset")
	rsTemp12.open sql,con,1
	if not rsTemp12.eof then
		GetNameOfCategoryLoai	=	rstemp12("TxtCategory")
	end if
    set rstemp12=nothing
    
End Function%>

<%Function GetProductOfCategoryLoai(CategoryLoai)
	sql12="Select * from CategoryLoai where CategoryLoai = '"& CategoryLoai &"'"
	Set rstemp12  = server.CreateObject("ADODB.Recordset")
	rsTemp12.open sql12,con,1
	if not rsTemp12.eof then
		GetProductOfCategoryLoai	=	Trim(rstemp12("isProduct"))
        set rstemp12=nothing
	end if
End Function%>


<%

function simalarString(strSournd,strSq)
	' tính theo phần trăm
	' điều kiện strSournd > strSq
	if InStr(strSournd,strSq)>0 then
		SimalarString = 100
		exit function
	end if
	iL2 = 	Len(strSq)
	iL1	=	Len(strSournd)
	
	iPT		=	iL2/iL1
	iPT		=	Round(iPT*100)
	iF iPT > 120 or iPT <25 or iL2 > 35 then
		simalarString	=	 0
		exit function
	end if

		
	strSournd	=	replace(strSournd,"  "," ")
	strSq	=	replace(strSq,"  "," ")
	arSournd	=	split(strSournd," ")
	arSq		=	split(strSq," ")
	iLen1	=	UBound(arSournd)
	iLen2	=	UBound(arSq)
	

	Set dictPercent = Server.CreateObject("Scripting.Dictionary")
	for i = 0 to iLen2
		if InStr(strSournd,arSq(i)) > 0 then			
			dictPercent.add i,100
		else
			tempPercent	= 0 
			for j = 0 to iLen1
				Word1	=	arSournd(j)
				Word2	=	arSq(i)
				subPercent	= simalarWord(Word1,Word2)	
				if tempPercent < subPercent then
					tempPercent = subPercent
				end if	
			next			
			dictPercent.add i,tempPercent
		end if	
	next
	tempPercent	= 0
	For Each Key in dictPercent
		tempPercent	=	tempPercent + dictPercent(Key)
	next
	simalarString = Round(tempPercent/dictPercent.count)
end function

function simalarWord(Word1,Word2)
	Word1	= 	trim(Word1)
	Word2	=	trim(Word2)
	if len(Word1) = 0 or len(Word2)=0 then
		simalarWord =  0 
		exit function
	end if
	if Len(Word1) < len(Word2) then
		strTemp = Word2
		Word2	=	Word1
		Word1	=	strTemp
	end if
	iK	=	Len(Word1) - len(Word2)
	if Word1=Word2 then
		simalarWord	= 100
		exit function
	elseif iK > 2 then
		simalarWord	= 0	
		exit function
	else
		subStr1	=	Word1
		subStr2	=	""
		for k = 1 to Len(Word2)
			charTemp= MID(Word2,k,1)
			iPos	=	instr(subStr1,charTemp)
			if iPos > 0 then
				subStr1	=	Left(subStr1,iPos-1) + Right(subStr1,Len(subStr1)-iPos)
			else
				subStr2	=	subStr2 + charTemp
			end if		
		next
		per1	=	Round(Len(subStr1)*100/len(Word1))
		per2	=	Round(Len(subStr2)*100/len(Word2))
		per		=	per1+per2
		if per > 100 then
			per	= 100
		end if
		simalarWord	= (100 - per)
	end if
end function
%>