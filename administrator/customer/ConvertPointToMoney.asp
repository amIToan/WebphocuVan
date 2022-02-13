<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%  
'	XSOFT
'   (C) Copyright XSOFT Corp. 2014
'   **************************
'   Cong ty phat trien cong nghe phan mem XSOFT  
'   Quan ly nhan su, ban hang, ton kho, tai chinh ke toan, tai chinh Price dinh.
'   Thiet ke website, thiet ke logo, catalog.
'   website:http://xsoft.com.vn
'   email:info@xsoft.com.vn  DT:04.2922.446
%>
<!--#include virtual="/include/config.asp"-->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/include/func_DateTime.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Chuyển điểm thành tiền</title>
<link href="css/styles.css" rel="stylesheet" type="text/css">

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	CMND		=	Session("CMND")
	sql = "SELECT SanPhamUser.SanPhamUser_ID, GetPoints FROM  SanPham_pay INNER JOIN SanPhamUser ON SanPham_pay.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID WHERE (SanPhamUser_Status = 2) and (GetPoints > 0) and (SanPhamUser.CMND ='"& CMND &"')"
	Set rsPoint=Server.CreateObject("ADODB.Recordset")
	rsPoint.open sql,con,1
	tPoint = 0
	if not rsPoint.eof then
		tPoint	=	0
		do while not rsPoint.eof 
			SanPhamUser_ID	=	rsPoint("SanPhamUser_ID")
			
			tPoint	=	rsPoint("GetPoints")
			sql	=	"update SanPham_pay set GetPoints = 0 where SanPhamUser_ID="&SanPhamUser_ID
			Set rsTemp=Server.CreateObject("ADODB.Recordset")
			rsTemp.open sql,con,1
			set rsTemp = nothing	
					
			rsPoint.movenext
		loop
		iniTK	=	tPoint*500
		if iniTK > 0 then
			Lydo	=	"Chuyển "&tPoint&" điểm thành tiền mặt"
			sql	=	"insert into TaiKhoan(CMND,iniTK,Lydo) values('"& CMND &"','"& iniTK &"',N'"& Lydo &"')"
			set rsTK = Server.CreateObject("ADODB.recordset")
			rsTK.open sql,con,1
			set rsTK	=	nothing
		end if
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"alert('Đã xử lý thành công!');"& vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline		
	else
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"alert('Đã xử lý không thành công xin hãy liên hệ với xbook!');"& vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline		
	end if
%>

</body>
</html>
