<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/func_tiny.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_cod")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
 iSoDH	=	GetNumeric(Request.Form("iSoTTAll"),0)
if iSoDH >= 0 then
	Ngay1=GetNumeric(Request.form("NgayCOD"),0)
	Thang1=GetNumeric(Request.form("ThangCOD"),0)
	Nam1=GetNumeric(Request.form("NamCOD"),0)
  	PayDate=Thang1 & "/" & Ngay1 & "/" & Nam1
	Response.Write(PayDate)
  	PayDate=FormatDatetime(PayDate)	
	for i = 0 to iSoDH 
		UserID	=	GetNumeric(Request.Form("User_ID_Chon"&i),0)
		if GetNumeric(Request.Form("ChonTT"&i),0) = 1 then			
			set rs=server.CreateObject("ADODB.Recordset")
			sql = "update SanPhamUser set NgayThanhToan = '"& PayDate &"' where SanPhamUser_ID=" & UserID
			rs.open sql,con,3
			set rs=nothing			
		end if
	next
end if
iSoDHHuy	=	GetNumeric(Request.Form("iSoHuyTTAll"),0)
if iSoDHHuy >= 0 then
	for i = 0 to iSoDHHuy 
		UserID	=	GetNumeric(Request.Form("User_ID_Huy"&i),0)
		if GetNumeric(Request.Form("HuyTT"&i),0) = 1 then
			set rs=server.CreateObject("ADODB.Recordset")
			sql = "update SanPhamUser set NgayThanhToan = '"& Null &"' where SanPhamUser_ID=" & UserID
			rs.open sql,con,3
			set rs=nothing
		end if
	next
end if
response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
		"window.close();"&vbNewline&_
	"//-->" & vbNewline &_
	"</script>" & vbNewline
%>

  
   