<!--#include virtual="/include/config.asp"-->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/include/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
		cmnd		=	request.Form("f_cmnd")
		fullname	=	request.Form("f_name")
		pass		=	Request.Form("f_pass")
		diachi		=	request.Form("f_diachi")
		quequan 	=	request.Form("f_quequan")
		email		=	request.Form("f_mail")
		tell		=	request.Form("f_tell")
		mobile		=	request.Form("f_mobile")
		strNgayCap	=	request.Form("dayCap")
		strThangCap	=	request.Form("morCap")
		strNamCap	=	request.Form("yearCap")	
		dateCap 	= 	strThangCap + "/"+ strNgayCap + "/" + strNamCap
		strNgaySinh	=	request.Form("f_NgaySinh")
		strThangSinh=	request.Form("f_ThangSinh")
		strNamSinh	=	request.Form("f_NamSinh")
		dateSinh 	=	strThangSinh + "/"+ strNgaySinh + "/" + strNamSinh
		noicap		=	request.Form("f_noicap")
		ProvinceID		=	Clng(request.Form("selTinh1"))
		DistrictID		=	Clng(request.Form("selHuyen1"))

		dim rsDiscuss
		set rsDiscuss=server.CreateObject("ADODB.Recordset")
		sql	=	"Update Account  Set "
		sql	=	sql & 	"Ngaycap 		= 	 '" & dateCap & "'"
		sql	=	sql & 	",Noicap		= 	N'" & noicap & "'"
		sql	=	sql & 	",Name 			= 	N'" & fullname & "'"
		sql	=	sql & 	",Ngaysinh 		= 	 '" & dateSinh & "'"			
		sql	=	sql & 	",nguyenquan	=	N'" & quequan & "'"
		sql	=	sql & 	",ProvinceID 		= 	"&ProvinceID
		sql	=	sql & 	",DistrictID 		= 	"&DistrictID
		sql	=	sql & 	",Diachi		=	N'" & diachi & "'"
		sql	=	sql & 	",Email		=	 '" & email & "'"
		sql	=	sql & 	",Tell		=	 '" & tell & "'"
		sql	=	sql & 	",Mobile		=	 '" & mobile & "'"
		sql	=	sql & 	",password	=	 '" & pass & "'"
		sql	=	sql	&	"Where CMND = '"& cmnd &"'"	
		rsDiscuss.open sql,con,1
		set rsDiscuss=nothing
		%>
<script language="javascript">
	window.close(); 
</script>
