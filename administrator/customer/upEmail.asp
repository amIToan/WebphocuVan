<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%	
	Action = Request.QueryString("Action")
	if Action="Del" then
		IDCare=Request.QueryString("IDCare")
		sql = "Delete CustomerCare where IDCare="&IDCare
		set rs=server.CreateObject("ADODB.Recordset")		
		rs.open sql,con,3
		set rs=nothing	
	%>
	<script language="javascript">
		alert('Đã thực hiện lệnh xóa!')
		window.history.back();
		window.location.reload();
	</script>
	<%				
	end if
	
	namek 		= 	trim(Request.Form("txtTen"))

	NgaySinh	=	Request.Form("NgaySinh")
	ThangSinh	=	Request.Form("ThangSinh")
	NamSinh		=	Request.Form("NamSinh")
	NgaySinh	=	ThangSinh + "/" + NgaySinh + "/" + NamSinh
	if isDate(NgaySinh)  = false  then
		NgaySinh = now()
	end if
	
	IDXungHo		=	GetNumeric(Request.Form("sel_xung_ho"),0)
	IDTamLy			=	GetNumeric(Request.Form("sel_tam_ly"),0)
	IDCongViec		=	GetNumeric(Request.Form("sel_cong_viec"),0)
	iDisable		=	TRim(Request.Form("iDisable"))
	dis_email 		=	1
	if iDisable	=	"false" then
		dis_email =	0
	else
		dis_email =	1
	end if
	
	DiaChi		=	Trim(Request.Form("txtDiaChi"))
	Tel			=	Trim(Request.Form("txtTel"))
	Email		=	trim(Request.Form("txtEmail"))
	GhiChu		=	Trim(Request.Form("txtGhichu"))
	
	addOrEddit 		= 	Request.Form("addOrEddit")
	set rs=server.CreateObject("ADODB.Recordset")
        if Email <> "" then
	if addOrEddit = 1 then
		ID	=	GetNumeric(Request.Form("ID"),0)
		sql = "Update Email set"
		sql	=	sql	+	" IDXungHo	='"& IDXungHo &"'"
		sql	=	sql	+	", NgaySinh	='"& NgaySinh &"'"		
		sql	=	sql	+	", Ten	=N'"& namek &"'"
		sql	=	sql	+	", IDTamLy	='"& IDTamLy &"'"
		sql	=	sql	+	", IDCongViec	=N'"& IDCongViec &"'"	
		sql	=	sql	+	", Diachi	=N'"& DiaChi &"'"
		sql	=	sql	+	", Dienthoai	=N'"& Tel &"'"
		sql	=	sql	+	", Email	='"& Email &"'"	
		sql	=	sql	+	", Disabled="&dis_email		
		sql	=	sql	+	", Ghichu	=N'"& GhiChu &"'"
		sql	=	sql	+	" Where ID=" & ID
		set rs=server.CreateObject("ADODB.Recordset")		
		rs.open sql,con,3
		set rs=nothing				
		
	else
		sql = 	"Insert into Email(IDXungHo,NgaySinh,Ten,IDTamLy,IDCongViec,Diachi,Dienthoai,email,Disabled,Ghichu) values("
		sql	=	sql	+ "'"& IDXungHo  &"'"
		sql	=	sql	+ ",'"& NgaySinh  &"'"
		sql	=	sql	+ ",N'"& namek  &"'"
		sql	=	sql	+ ",'"& IDTamLy  &"'"
		sql	=	sql	+ ",'"& IDCongViec  &"'"
		sql	=	sql	+ ",N'"& DiaChi  &"'"
		sql	=	sql	+ ",'"& Tel  &"'"
		sql	=	sql	+ ",'"& Email  &"'"
		sql	=	sql	+	",'"&dis_email&"'"			
		sql	=	sql	+ ",N'"& GhiChu  &"')"	
		set rs=server.CreateObject("ADODB.Recordset")		
		rs.open sql,con,3
		set rs=nothing				
	end if
end if
	

%>		
	<script language="javascript">
		window.opener.location.reload();
		window.close();
	</script>