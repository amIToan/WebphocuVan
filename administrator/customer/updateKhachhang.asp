<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>

<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td ></td>
  </tr>
  <tr>
  	<td ><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td class="author">
		  <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fKhachHang">
		  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center">&nbsp;</td>
            </tr>
            <tr>
              <td align="center"><span class="CTxtContent">Cập nhật vào nhóm:
                  <select name="selNhomEMail" id="selNhomEMail">
                      <option value="0" selected="selected">&nbsp;&nbsp;&nbsp;&nbsp;</option>
                      <%
  	sql = "Select * From EmailNhom ORDER BY TenNhom "
  	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	do while not rs.eof 
%>
                      <option value="<%=rs("IDNhomEmail")%>" <%if rs("IDNhomEmail") = IDNhom then Response.Write("selected")%>><%=rs("TenNhom")%></option>
                      <%		rs.movenext
	loop
  set rs = nothing
%>
                    </select>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="submit" name="Submit2" value=" Cập nhận email ">
				<input type="hidden" name="up" value="1">
              </span></td>
            </tr>
          </table>
		  </form>
		  </td>
        </tr>
        <tr>
          <td class="author">&nbsp;</td>
        </tr>
<%
	selNhomEMail	=	Request.Form("selNhomEMail")
	up	=	GetNumeric(Request.Form("up"),0)
	if up = 1 then
%>		
        <tr>
          <td class="author"><div align="center">DANH SÁCH KHÁCH HÀNG MỚI CẬP NHẬT </div></td>
        </tr>
        <tr>
          <td>
<%

	sql = "Select * from SanPhamUser"
	set rs	=	server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	stt = 0
	if not rs.eof then
	do while not rs.eof
		if isExistEmail(rs("SanPhamUser_Email")) = 0 then
			stt = stt + 1
			Names 	= Trim(rs("SanPhamUser_Name"))
			Xungho	=	""
			iXH	=	false
			if  Instr(UCase(Names),"ANH")=1 then
				Xungho	=	"Anh"
				Names =right(Names,Len(Names)-3)
				
			elseif(Instr(UCase(Names),"CHỊ")>0)  then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Chị","")
				
			elseif Instr(UCase(Names),"VĂN") > 0  then
				Xungho	=	"Anh"
				
			elseif(Instr(UCase(Names),"THỊ")>0) then
				Xungho	=	"Chị"
			elseif(Instr(Ucase(Names),"MS")>0) then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Ms","")	
													
			elseif(Instr(Ucase(Names),"MR")>0) then
				Xungho	=	"Anh"
				Names	=	replace(Names,"Mr","")	
								
			elseif(Instr(Names,"Bạn")=1) or (Instr(Names,"bạn")=1) then
				Xungho	=	"Bạn"				
				Names	=	Names - right(Names,Len(Names)-3)																								
			end if
			NgaySinh	=	rs("SanPhamUser_Date")
			if isdate(NgaySinh) then
				NgaySinh	= now
			end if
			sql = 	"insert into Email(Ten,NgaySinh,Diachi,DienThoai,Email) "
			sql =	sql +" values(N'"&Names&"','"& NgaySinh &"'"
			sql =	sql +" ,N'"& rs("SanPhamUser_Address") &"',N'"& rs("SanPhamUser_Tell") &"'"
			sql =	sql +" ,N'"& rs("SanPhamUser_Email") &"')"
			set rsTemp	=	server.CreateObject("ADODB.Recordset")
			on error Resume next
			rsTemp.open sql,con,1
			Response.Write(stt&". "&Names&" - "&rs("SanPhamUser_Email")&"<br>")
			Set rsTemp	=	nothing	
		end if					
		rs.movenext
	loop
	set rs	= nothing
	else
		Response.Write("Không có khách hàng mới trong đơn hàng")
	end if
	
	' Cập nhật trong tài khoản khách hàng
	sql = "Select * from Account"
	set rs	=	server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	stt = 0
	if not rs.eof then
	do while not rs.eof
		if isExistEmail(rs("Email")) = 0 then
			stt = stt + 1
			Names 	= Trim(rs("Name"))
			Xungho	=	""
			if  Instr(UCase(Names),"ANH")=1 then
				Xungho	=	"Anh"
				Names =right(Names,Len(Names)-3)
				
			elseif(Instr(UCase(Names),"CHỊ")>0)  then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Chị","")
				
			elseif Instr(UCase(Names),"VĂN") > 0  then
				Xungho	=	"Anh"
				
			elseif(Instr(UCase(Names),"THỊ")>0) then
				Xungho	=	"Chị"
			elseif(Instr(Ucase(Names),"MS")>0) then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Ms","")	
													
			elseif(Instr(Ucase(Names),"MR")>0) then
				Xungho	=	"Anh"
				Names	=	replace(Names,"Mr","")	
								
			elseif(Instr(Names,"Bạn")=1) or (Instr(Names,"bạn")=1) then
				Xungho	=	"Bạn"				
				Names	=	Names - right(Names,Len(Names)-3)																								
			end if
			sql = 	"insert into Email(Ten,NgaySinh,Diachi,DienThoai,Email) "
			sql =	sql +" values(N'"&Names&"','"& rs("NgaySinh") &"'"
			sql =	sql +" ,N'"& rs("diachi") &"',N'"& rs("mobile") &"'"
			sql =	sql +" ,N'"& rs("Email") &"')"
			set rsTemp	=	server.CreateObject("ADODB.Recordset")
			on error Resume next
			rsTemp.open sql,con,1
			Response.Write(stt&". "&Names&" - "&rs("Email")&"<br>")
			Set rsTemp	=	nothing	
		end if
		rs.movenext
	loop
	set rs	= nothing
	else
		Response.Write("<br>Không có khách hàng mới trong tài khoản")
	end if	
	
	
	' Cập nhật trong Ý kiến
	sql = "Select * from Comment"
	set rs	=	server.CreateObject("ADODB.Recordset")
	rs.open sql,con,3
	stt = 0
	if not rs.eof then
	do while not rs.eof
		if isExistEmail(rs("email")) = 0 then
			stt = stt + 1
			Names 	= Trim(rs("hovaten"))
			Xungho	=	""
			if  Instr(UCase(Names),"ANH")=1 then
				Xungho	=	"Anh"
				Names =right(Names,Len(Names)-3)
				
			elseif(Instr(UCase(Names),"CHỊ")>0)  then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Chị","")
				
			elseif Instr(UCase(Names),"VĂN") > 0  then
				Xungho	=	"Anh"
				
			elseif(Instr(UCase(Names),"THỊ")>0) then
				Xungho	=	"Chị"
			elseif(Instr(Ucase(Names),"MS")>0) then
				Xungho	=	"Chị"
				Names	=	replace(Names,"Ms","")	
													
			elseif(Instr(Ucase(Names),"MR")>0) then
				Xungho	=	"Anh"
				Names	=	replace(Names,"Mr","")	
								
			elseif(Instr(Names,"Bạn")=1) or (Instr(Names,"bạn")=1) then
				Xungho	=	"Bạn"				
				Names	=	Names - right(Names,Len(Names)-3)																								
			end if
			sql = 	"insert into Email(Ten,NgaySinh,DienThoai,Email) "
			sql =	sql +" values(N'"&Names&"','"& rs("ngaytao") &"'"
			sql =	sql +" ,N'"& rs("tel") &"'"
			sql =	sql +" ,N'"& rs("email") &"')"
			set rsTemp	=	server.CreateObject("ADODB.Recordset")
			on error Resume next
			rsTemp.open sql,con,1
			Response.Write(stt&". "&Names&" - "&rs("Email")&"<br>")
			Set rsTemp	=	nothing	
		end if
		rs.movenext
	loop
	set rs	= nothing
	else
		Response.Write("<br>Không có khách hàng mới trong tài khoản")
	end if	
	
%>

</td>
        </tr>
<%end if%>		
      </table>

	  </td>
  </tr>
    <td ></td>
  </tr>
</table>
</body>
</html>