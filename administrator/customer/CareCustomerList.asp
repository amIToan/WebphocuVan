<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%IF Request.form("action")="Search" then
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	iTim =GetNumeric(Request.form("cbAll"),0)	
	strDieuKien	=	Request.Form("txtDieuKien")
	iDieuKien	=	GetNumeric(Request.Form("selDieuKien"),0)
	iSapXep			=	GetNumeric(Request.Form("selSapXep"),0)
	iTangPricem		=	GetNumeric(Request.Form("raTangPricem"),0)
	
	NameProvince		=	Trim(Request.Form("selTinh1"))
	fTien1	=	Chuan_money(Request.Form("txtTien1"))
	cbAll	=	GetNumeric(Request.Form("cbAll"),0)
	m_day	=	GetNumeric(Request.Form("m_day"),0)
	stt		=	GetNumeric(Request.Form("stt"),0)
	
	iTop	=	GetNumeric(Request.Form("txtTop"),0)
	iTopBegin	=	GetNumeric(Request.Form("txtTopBegin"),0)
	
	iCheckThuTay	=	GetNumeric(Request.Form("iCheckThuTay"),0)
	act		=	"SendMailTT.asp"
ELSE
	Day1 = now() - 30
	Ngay1=Day(Day1)
	Thang1=Month(Day1)
	Nam1=Year(Day1)
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
	iDieuKien = 0
	cbAll	=	0
	act		=	"List_email.asp"
END IF
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/new_mail_accept.gif"
	Title_This_Page="Khách hàng -> Danh sách email khách hàng"
	Call header()
	Call Menu()
	
	
%>

<form name="fEmail" method="post" action="<%=act%>" >
<%IF Request.form("action")<>"Search" then%>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td >
	<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" style="border:#CCCCCC solid 1px;">
      <tr>
        <td align="center">
		<table width="80%" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
        <tr>
          <td colspan="4" align="right" valign="middle" ><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
            <%
			Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call List_Month_WithName(Thang1,"MM","Thang1")
			Call  List_Year_WithName(Nam1,"YYYY",1960,"Nam1")
		%>
            <img src="../images/right.jpg" width="9" height="9" align="absmiddle">
            <%
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call List_Month_WithName(Thang2,"MM","Thang2")
			Call  List_Year_WithName(Nam2,"YYYY",1960,"Nam2")
		%>
			</div></td>
          </tr>
        
        <tr>
          <td width="15%" align="center" valign="middle" class="CTxtContent" ><div align="right"><em>Điều kiện:
            
          </em></div></td>
          <td width="31%" align="center" valign="middle" class="CTxtContent" ><div align="left"><em>
            <input name="txtDieuKien" type="text" id="txtDieuKien" value="<%=strDieuKien%>">
          </em></div></td>
          <td width="20%" align="center" valign="middle" class="CTxtContent" ><div align="right"><em>Tìm theo:</em></div></td>
          <td width="34%" align="center" valign="middle" class="CTxtContent" ><div align="left">
            <select name="selDieuKien" id="selDieuKien">
              <option value="0" selected>CSKH</option>
              <option value="1" <%if iDieuKien = 1 then%>selected<%end if%>>Tên khách</option>
              <option value="2" <%if iDieuKien = 2 then%>selected<%end if%>>Email</option>
              <option value="3" <%if iDieuKien = 3 then%>selected<%end if%>>Điện thoại</option>
            </select>
          </div></td>
        </tr>
        <tr>
          <td colspan="4" valign="middle" class="CTxtContent"  align="center">Hiển thị từ :
            <input name="txtTopBegin" type="text" size="5" maxlength="10" value="0" onBlur="javascript: checkIsNumber(this)">
            Đến:
            <input name="txtTop" type="text" size="5" maxlength="10" value="100" onBlur="javascript: checkIsNumber(this)"> 
            khách hàng </td>
          </tr>
        
  
        <tr>
          <td align="center" valign="middle" class="CTxtContent" ><div align="right">Sắp xếp:            </div></td>
          <td align="center" valign="middle" class="CTxtContent" ><div align="left">
            <select name="selSapXep" id="selSapXep">
              <option value="0" selected <%if iSapXep = 0 then%>selected<%end if%>></option>
              <option value="1" <%if iSapXep = 1 then%>selected<%end if%>>Họ và tên</option>
              <option value="2" <%if iSapXep = 2 then%>selected<%end if%>>Số lần đặt mua</option>
              <option value="3" <%if iSapXep = 3 then%>selected<%end if%>>Tổng tiền mua hàng</option>
              <option value="4">Ngày đặt mua đầu</option>
            </select>
          </div></td>
          <td colspan="2" align="center" valign="middle" class="CTxtContent" ><input name="raTangPricem" type="radio" value="1" checked <%if iTangPricem = 1 then Response.Write("checked") end if %>>
            Giảm dần
              <input name="raTangPricem" type="radio" value="2" <%if iTangPricem = 2 then Response.Write("checked") end if %>> 
            Tăng dần				  </td>
          </tr>
        <tr>
          <td colspan="4" align="center" valign="middle" >
           <input type="hidden" name="action" value="Search">            <input type="submit" name="Submit11" value="  Xem " >          </td>
          </tr>
      </table>		</td>
      </tr>
    </table>
<%end if%>	
	<br>
	<br>
<%IF Request.form("action")="Search" then

%>
	<table <%if iCheckThuTay = 0 then%> width="95%" <%else Response.Write("width=""500""") end if%> border="0" align="center" cellpadding="1" cellspacing="1">
<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
		
	sql="select Top "& iTop &" * from Email"
	if cbAll <> 1 then
	if iTim <> 1 then
		select case iDieuKien
			case 0 
				sql = sql & " WHERE "		
			case 1 
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Ten)} like N'%" & strDieuKien & "%' and "
			case 2
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Email)} like N'%" & strDieuKien & "%' and "		
			case 3
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where (Dienthoai like '%" & strDieuKien & "%') and "			
		end select
	else
		sql = sql & " WHERE "		
	end if
	
	if m_day = 0 then	
		sql=sql+"  (DATEDIFF(dd,NgaySinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgaySinh,'" & ToDate &"') >= 0) "
	else
		sql=sql+"  (DATEDIFF(dd,CreateDate,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,CreateDate,'" & ToDate &"') >= 0) " 
	end if
	
	if iCheckNhom > 0 then
		ik	=1
		IDNhom	=	GetNumeric(Request.Form("IDEmail"&ik),0)
		sql=sql+" AND (IDXungHo = '"& IDNhom &"' or IDTamLy = '"& IDNhom &"' or IDCongViec = '"& IDNhom &"')  "
		for ik = 2 to stt
			IDNhom	=	GetNumeric(Request.Form("IDEmail"&ik),0)
			if IDNhom > 0 then
				sql=sql+" or (IDXungHo = '"& IDNhom &"' or IDTamLy = '"& IDNhom &"' or IDCongViec = '"& IDNhom &"')  "
			end if
		next		
	end if
	if NameProvince <> "not" then
		sql=sql+" and Diachi like N'%" & NameProvince & "%'"
	end if
	end if

	Select case iSapXep
		case 1
			sql=sql & " order by Ten"
			if iTangPricem = 1 then
				sql = sql & " DESC "
			end if
		case 4
			sql=sql & " order by NgaySinh"
			if iTangPricem = 1 then
				sql = sql & " DESC "
			end if				
	end select
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,3
  length = rs.recordcount-1-iTopBegin
  Redim arEmail(length,9)
  i = 0
  j	= 0
  Do while not rs.eof
  	if j >= iTopBegin then
		ID		=	rs("ID")
		Xungho	=	Get_Name_Group(rs("IDXungho"))
		Ten		=	rs("ten")
		Tel		=	rs("DienThoai")
		Email	=	rs("Email")
		Diachi	=	rs("Diachi")
		
		sql	=	"Select SanPhamUser_ID,CMND,SanPhamUser_Status From SanPhamUser where SanPhamUser_Email=N'"& Email &"'"
		Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
		rsTemp1.open sql,Con,3
		TTien	=	0
		iSoHD	=	0
		iDHMoi	=	0
		iDHDangXL	=	0
		iDHHuy		=	0
		do while not rsTemp1.eof		
			
			select case rsTemp1("SanPhamUser_Status")
				case 0
					iDHMoi = iDHMoi + 1
				case 1,4,5,8,7,8
					iDHDangXL=iDHDangXL+1
				case 2	
					TTien	 =	 TTien + LamTronTien(TongTienTrenDonHang(rsTemp1("SanPhamUser_ID"),rsTemp1("CMND")))
					iSoHD = iSoHD + 1
				case 3
					iDHHuy=iDHHuy+1					
			end select	
			rsTemp1.movenext
		loop
		Set rsTemp1 = nothing	
		arEmail(i,0)	=	ID
		arEmail(i,1)	=	Ten
		arEmail(i,2)	=	Tel
		arEmail(i,3)	=	Email
		arEmail(i,4)	=	Diachi
		arEmail(i,5) 	= 	iSoHD
		arEmail(i,6) 	= 	TTien
		arEmail(i,7) 	= 	iDHMoi
		arEmail(i,8) 	= 	iDHDangXL
		arEmail(i,9) 	= 	iDHHuy
		i=i+1
	end if
	j=j+1	
  	rs.movenext
  loop
%>

<tr>
  <td colspan="8" class="CTxtContent">Hiện có <b><%=length%></b> Khách hàng.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    </td>
  </tr>
<tr>
	<td width="35" height="19" align="center" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
	<%if iCheckThuTay = 0 then%>	
	<td width="49" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>" align="center">
	  Chọn	</td><%end if%>
	<td width="151" bgcolor="#FFFFCC" class="CFontVerdana10" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Họ và tên</td>
	<%if iCheckThuTay = 0 then%>
	<td width="308" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Email</td>
	<td width="183" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Nhóm  </td>
	<td width="67" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Tiền/ĐH</td>
	<td width="51" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
	<td width="34" align="center" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
	<%end if%>
</tr>
<%
  i = 0
  redim arTemp(9)
Select case iSapXep
	case 3
		For i = 0 to  length
			for j = i to length
				if arEmail(i,6) < arEmail(j,6) then
					arTemp(0) = arEmail(i,0)
					arTemp(1) = arEmail(i,1)
					arTemp(2) = arEmail(i,2)
					arTemp(3) = arEmail(i,3)
					arTemp(4) = arEmail(i,4)
					arTemp(5) = arEmail(i,5)
					arTemp(6) = arEmail(i,6)
					arTemp(7) = arEmail(i,7)
					arTemp(8) = arEmail(i,8)
					arTemp(9) = arEmail(i,9)
						
					arEmail(i,0) = arEmail(j,0)
					arEmail(i,1) = arEmail(j,1)
					arEmail(i,2) = arEmail(j,2)
					arEmail(i,3) = arEmail(j,3)
					arEmail(i,4) = arEmail(j,4)
					arEmail(i,5) = arEmail(j,5)
					arEmail(i,6) = arEmail(j,6)
					arEmail(i,7) = arEmail(j,7)
					arEmail(i,8) = arEmail(j,8)
					arEmail(i,9) = arEmail(j,9)

					arEmail(j,0) = arTemp(0)
					arEmail(j,1) = arTemp(1)
					arEmail(j,2) = arTemp(2)
					arEmail(j,3) = arTemp(3)
					arEmail(j,4) = arTemp(4)
					arEmail(j,5) = arTemp(5)
					arEmail(j,6) = arTemp(6)
					arEmail(j,7) = arTemp(7)
					arEmail(j,8) = arTemp(8)
					arEmail(j,9) = arTemp(9)

				end if
			next					
		next
	case 2
		For i = 0 to  length
			for j = i to length
				if arEmail(i,5) < arEmail(j,5) then
					arTemp(0) = arEmail(i,0)
					arTemp(1) = arEmail(i,1)
					arTemp(2) = arEmail(i,2)
					arTemp(3) = arEmail(i,3)
					arTemp(4) = arEmail(i,4)
					arTemp(5) = arEmail(i,5)
					arTemp(6) = arEmail(i,6)
					arTemp(7) = arEmail(i,7)
					arTemp(8) = arEmail(i,8)
					arTemp(9) = arEmail(i,9)
						
					arEmail(i,0) = arEmail(j,0)
					arEmail(i,1) = arEmail(j,1)
					arEmail(i,2) = arEmail(j,2)
					arEmail(i,3) = arEmail(j,3)
					arEmail(i,4) = arEmail(j,4)
					arEmail(i,5) = arEmail(j,5)
					arEmail(i,6) = arEmail(j,6)
					arEmail(i,7) = arEmail(j,7)
					arEmail(i,8) = arEmail(j,8)
					arEmail(i,9) = arEmail(j,9)

					arEmail(j,0) = arTemp(0)
					arEmail(j,1) = arTemp(1)
					arEmail(j,2) = arTemp(2)
					arEmail(j,3) = arTemp(3)
					arEmail(j,4) = arTemp(4)
					arEmail(j,5) = arTemp(5)
					arEmail(j,6) = arTemp(6)
					arEmail(j,7) = arTemp(7)
					arEmail(j,8) = arTemp(8)
					arEmail(j,9) = arTemp(9)
				end if
			next					
		next
	end select
	iBegin = 0
	iKetThuc = length
	iStep 	 = 1
	if (iTangPricem <> 1  and (iSapXep = 2 or iSapXep = 3)) then
		iBegin = length
		iKetThuc = 0
		iStep = -1
	end if		
  For i = iBegin to  iKetThuc Step iStep
  		m_xungho	=	Get_Name_Group(arEmail(i,0))
		%>
		<tr <%if i mod 2=0 and iCheckThuTay = 0  then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
			<td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=i+iTopBegin+1%></td>
			<%if iCheckThuTay = 0 then%>
			<td align="center" style="<%=setStyleBorder(1,1,0,1)%>">
			<input type="checkbox" name="CbEmailKhach<%=i%>" value="1">	<br>		
			<input type="hidden" name="IDEmail<%=i%>" value="<%=arEmail(i,0)%>">			</td>
			
			<%end if%>
			<td class="CTxtContent" style="<%=setStyleBorder(0,1,0,1)%>">
			<%if iCheckThuTay = 0 then%>
			<%=m_xungho&" "%><b><%=arEmail(i,1)%></b>
			<br>
			<font class="CSubTitle"><u>Tel</u>:<%=arEmail(i,2)%></font>
			<%else%>
				<%=m_xungho&" "%> <b><%=arEmail(i,1)%></b><br>
				Đ/C:<%=arEmail(i,4)%><br>
				ĐT: <%=arEmail(i,2)%>
			<%end if%>	
			</td>
			<%if iCheckThuTay = 0 then%>
			<td style="<%=setStyleBorder(0,1,0,1)%>" >
			<font class="CTxtContent"><%=arEmail(i,4)%></font><br>
			<a href="send_mail.asp?email=<%=arEmail(i,5)%>" class="CSubTitle"><%=arEmail(i,3)%></a>
			<input type="hidden" name="hEmail<%=i%>" value="<%=arEmail(i,3)%>">
			<br>              </td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
				<%=Get_Name_Group(Get_IDGroup_FromEmail(arEmail(i,3),"IDTamly"))%><br>
				<%=Get_Name_Group(Get_IDGroup_FromEmail(arEmail(i,3),"IDCongViec"))%>&nbsp;</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
			<%=Dis_str_money(arEmail(i,6))%>/ <%=arEmail(i,5)%><br>
			<font class="CSubTitle">
			<%
				if 	arEmail(i,7) > 0 then
					Response.Write("Mới: "&arEmail(i,7)&"<br>")
				end if
				
				if arEmail(i,8) > 0 then
					Response.Write("Đang xử lý: "&arEmail(i,8)&"<br>")		
				end if
				
				if arEmail(i,9) > 0 then 
					Response.Write("Hủy: "&arEmail(i,9)&"<br>")		
				end if
			%></font>			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><a href="javascript: winpopup('History_email.asp','<%=arEmail(i,3)%>&Name=<%=m_xungho+" "+arEmail(i,1)%>',990,600);" class="CSubMenu">Lịch sử</a></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
			<a href="updateEmail.asp?addOrEddit=1&ID=<%=arEmail(i,0)%>">
			<img src="../../images/icons/article.gif" width="16" height="16" border="0" align="absmiddle"></a>
			<%if Session("iQuanTri") = 1 then %>
			<img src="../../images/icons/icon_pmdead.gif" border="0" align="absmiddle" onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa nhân viên này không?'); if(yn) {window.location = 'delEmail.asp?ID=<%=arEmail(i,0)%>'}" >
			
			<%end if%>			</td>
			<%end if%>
		</tr>
	<%
	next
	%>
	<%if iCheckThuTay = 0 then%>
	<tr>
		<td>&nbsp;</td>
		<td style="<%=setStyleBorder(1,1,0,1)%>" align="center">
		<%
		iSoEmail	=	i
		%>
		<input type="hidden" name="iSoEmail" value="<%=iSoEmail%>">
		<input type="checkbox" name="CbAllEmail" value="1" onClick="javascript:OnCheckAll()" >		</td>
		<td><input type="submit" name="Submit3" value=" Gửi lựa chọn" onClick="javascript: ReSub();"></td>
		<td><input type="button" name="Submit32" value=" Gửi tất cả" onClick="javascript: window.location = 'SendMailTT.asp?All=ok'"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>	
		<td align="center">			</td>			
	</tr>
	<%end if%>
</table>
<%end if%>
	<br></td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
</form>

</body>
</html>
<script language="javascript">
function OnCheckAll()
{
	iNumSP	=	document.fEmail.iSoEmail.value-1;	
	if (document.fEmail.CbAllEmail.checked == true)
		iCbAll = 1
	else
		iCbAll = 0
	
	for(jj=0;jj<=iNumSP;jj++)
	{
		if (iCbAll == 1)
			str = "document.fEmail.CbEmailKhach"+jj+".checked = true";
		else
			str = "document.fEmail.CbEmailKhach"+jj+".checked = false";
		eval(str);	
	}
	return;
	
}

function ShowEmailGroup()
{
	if (document.fEmail.CheckGroupEmail.checked == true)
		document.getElementById("ShowGroupEmail").style.display="";
	else 
		document.getElementById("ShowGroupEmail").style.display="none";
		

}
</script>

