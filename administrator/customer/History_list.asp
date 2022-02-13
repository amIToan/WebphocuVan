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
	raNgay			=	GetNumeric(Request.Form("raNgay"),0)
	iSapXep			=	GetNumeric(Request.Form("selSapXep"),0)
	iTangPricem		=	GetNumeric(Request.Form("raTangPricem"),0)
ELSE
	Day1 = now() - 30
	Ngay1=Day(Day1)
	Thang1=Month(Day1)
	Nam1=Year(Day1)
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
	iDieuKien = 1
	raNgay = 1
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
	img	=	"../../images/icons/icon_customer.gif"
	Title_This_Page="Khách hàng -> Lịch sử giao dịch"
	Call header()
	Call Menu()
	
	
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	    <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fKhachHang" onSubmit="return checkme();">	
	  <table width="95%" align="center" cellpadding="2" cellspacing="2">
        <tr>
          <td align="right" valign="middle" ><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
                  <%
			Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call List_Month_WithName(Thang1,"MM","Thang1")
			Call  List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
		%>
                  <img src="../images/right.jpg" width="9" height="9" align="absmiddle">
                  <%
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call List_Month_WithName(Thang2,"MM","Thang2")
			Call  List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
		%>
          </div></td>
        </tr>
        <tr>
          <td align="center" valign="middle" ><input name="raNgay" type="radio" value="1" checked <%if iTangPricem = 1 then Response.Write("checked") end if %>>
            Ngày đăng ký  
              <input name="raNgay" type="radio" value="2" <%if iTangPricem = 2 then Response.Write("checked") end if %>>
            Ngày truy cập lần cuối </td>
        </tr>
        <tr>
          <td align="center" valign="middle" class="CTxtContent" ><em>Điều kiện: </em>
              <input name="txtDieuKien" type="text" id="txtDieuKien" value="<%=strDieuKien%>">
              <em>Tìm theo:</em>
              <select name="selDieuKien" id="selDieuKien">
                <option value="0" selected <%if iDieuKien = 0 then%>selected<%end if%>></option>
                <option value="1" <%if iDieuKien = 1 then%>selected<%end if%>>Tên khách</option>
                <option value="2" <%if iDieuKien = 2 then%>selected<%end if%>>Email</option>
                <option value="3" <%if iDieuKien = 3 then%>selected<%end if%>>Điện thoại</option>
                <option value="4" <%if iDieuKien = 4 then%>selected<%end if%>>CMND</option>
              </select>          </td>
        </tr>
        <tr>
          <td align="center" valign="middle" class="CTxtContent" >&nbsp;</td>
        </tr>
        <tr>
          <td align="center" valign="middle" class="CTxtContent" >Sắp xếp: 
            <select name="selSapXep" id="selSapXep">
              <option value="0" selected <%if iSapXep = 0 then%>selected<%end if%>></option>
              <option value="1" <%if iSapXep = 1 then%>selected<%end if%>>Họ và tên</option>
              <option value="2" <%if iSapXep = 2 then%>selected<%end if%>>Số lần đặt mua</option>
              <option value="3" <%if iSapXep = 3 then%>selected<%end if%>>Tổng tiền mua hàng</option>
              <option value="4" <%if iSapXep = 4 then%>selected<%end if%>>Ngày đăng ký</option>
              </select>
            <input name="raTangPricem" type="radio" value="1" checked <%if iTangPricem = 1 then Response.Write("checked") end if %>>
            Giảm dần
            <input name="raTangPricem" type="radio" value="2" <%if iTangPricem = 2 then Response.Write("checked") end if %>> 
            Tăng dần</td>
        </tr>
        <tr>
          <td align="center" valign="middle" width="100%" >
              <input name="cbAll" type="checkbox" id="cbAll" value="1">
            Tìm tất cả 
            <input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
            <input type="hidden" name="action" value="Search">          </td>
        </tr>
      </table>
	  </form>
	  </td>
    </tr>
    <tr>
      <td background="../../images/T3.jpg" height="8"></td>
    </tr>
</table>
<br>
<%
IF Request.form("action")="Search"  THEN
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)

	Dim rs
	Dim emailx
	set rs=server.CreateObject("ADODB.Recordset")
	sql="select * from Account"
	if iTim <> 1 then
		select case iDieuKien
			case 0 
				sql = sql & " WHERE "		
			case 1 
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Name)} like N'%" & strDieuKien & "%' and "
			case 2
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Email)} like N'%" & strDieuKien & "%' and "		
			case 3
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where (Tell like '%" & strDieuKien & "%' or mobile like '%" & strDieuKien & "%') and "		
			case 4
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where CMND like '%" & strDieuKien & "%' and "	
		end select
	else
		sql = sql & " WHERE "		
	end if	
		
		if raNgay = 1 then
			sql=sql+"  (DATEDIFF(dd,CreationDate,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,CreationDate,'" & ToDate &"') >= 0) " 
		else
			sql=sql+"  (DATEDIFF(dd,LastLoginDate,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,LastLoginDate,'" & ToDate &"') >= 0) " 		
		end if
		
		Select case iSapXep
			case 1
				sql=sql & " order by Name"
				if iTangPricem = 1 then
					sql = sql & " DESC "
				end if
			case 4
				sql=sql & " order by CreationDate"
				if iTangPricem = 1 then
					sql = sql & " DESC "
				end if				
		end select

	rs.open sql,con,1
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</font></strong></td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	else
	  length = rs.recordcount - 1
	  Redim arCMT(length,9)
	  i = 0
	  Do while not rs.eof
			CMT		=	rs("CMND")
			Ho_Ten	= 	rs("Name")
			CreationDate=rs("CreationDate")
			DiaChi	=	rs("diachi")
			if GetNameDistrict(rs("DistrictID")) <> "" then
				DiaChi = DiaChi + "; " + GetNameDistrict(rs("DistrictID"))
			end if
			
			if GetNameProvince(rs("ProvinceID")) <> "" then
				DiaChi = DiaChi + "; " + GetNameProvince(rs("ProvinceID"))
			end if
 	
			if rs("Tell") <> "" and rs("mobile") <> "" then
				DienThoai= 	rs("Tell") + " / "+rs("mobile")
			else
				DienThoai= 	Trim(rs("Tell") + rs("mobile"))
			end if	
		
			sql 	= 	"SELECT Account.CMND,SanPhamUser_Name,SanPhamUser_ID,SanPhamUser_Address,SanPhamUser_Tell,SanPhamUser_Status FROM Account INNER JOIN SanPhamUser ON Account.CMND = SanPhamUser.CMND" 
			sql 	= 	sql + " WHERE(Account.CMND = '"& CMT &"')"
			set rs1=server.CreateObject("ADODB.Recordset")
			rs1.open sql,con,1
			iSoHD = 0
			TTien=0
			iDHMoi = 0
			iDHDangXL = 0
			iDHOK = 0
			iDHHuy= 0
			Do while not rs1.eof
				select case rs1("SanPhamUser_Status")
					case 0
						iDHMoi = iDHMoi + 1
					case 1,4,5,6,7,8
						iDHDangXL=iDHDangXL+1
					case 2
						iTien	= 	LamTronTien(TongTienTrenDonHang(rs1("SanPhamUser_ID"),CMT))	
						TTien	 =	 TTien + iTien
						iSoHD = iSoHD + 1
					case 3
						iDHHuy=iDHHuy+1					
				end select				
				rs1.movenext
			loop
			arCMT(i,0) = CMT
			arCMT(i,1) = Ho_Ten
			arCMT(i,2) = DiaChi
			arCMT(i,3) = DienThoai
			arCMT(i,4) = CreationDate
			arCMT(i,5) = iSoHD
			arCMT(i,6) = TTien
			arCMT(i,7) = iDHMoi
			arCMT(i,8) = iDHDangXL
			arCMT(i,9) = iDHHuy
			i = i +1	
			set rs1	= nothing
		rs.movenext
	  loop
	  rs.close
	set rs=nothing
	
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr align="center" bgcolor="FFFFFF">
    <td width="51" bgcolor="#FFCC00" style="<%=setStyleBorder(1,1,1,1)%>"><strong>TT</strong></td>
    <td width="165" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Họ tên </strong></td>
    <td width="340" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Địa chỉ </strong></td>
    <td width="127" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tel</strong></td>
    <td width="73" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày đăng ký </strong></td>
    <td width="37" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Đã xlý </strong></td>
    <td width="70" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tổng tiền </strong></td>
    <td width="88" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tình trạng khác </strong></td>
    <td width="44" bgcolor="#FFCC00" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
  </tr>
  
  <%
  i = 0
  redim arTemp(9)
Select case iSapXep
	case 3
		For i = 0 to  length
			for j = i to length
				if arCMT(i,6) < arCMT(j,6) then
					arTemp(0) = arCMT(i,0)
					arTemp(1) = arCMT(i,1)
					arTemp(2) = arCMT(i,2)
					arTemp(3) = arCMT(i,3)
					arTemp(4) = arCMT(i,4)
					arTemp(5) = arCMT(i,5)
					arTemp(6) = arCMT(i,6)
					arTemp(7) = arCMT(i,7)
					arTemp(8) = arCMT(i,8)
					arTemp(9) = arCMT(i,9)
						
					arCMT(i,0) = arCMT(j,0)
					arCMT(i,1) = arCMT(j,1)
					arCMT(i,2) = arCMT(j,2)
					arCMT(i,3) = arCMT(j,3)
					arCMT(i,4) = arCMT(j,4)
					arCMT(i,5) = arCMT(j,5)
					arCMT(i,6) = arCMT(j,6)
					arCMT(i,7) = arCMT(j,7)
					arCMT(i,8) = arCMT(j,8)
					arCMT(i,9) = arCMT(j,9)

					arCMT(j,0) = arTemp(0)
					arCMT(j,1) = arTemp(1)
					arCMT(j,2) = arTemp(2)
					arCMT(j,3) = arTemp(3)
					arCMT(j,4) = arTemp(4)
					arCMT(j,5) = arTemp(5)
					arCMT(j,6) = arTemp(6)
					arCMT(j,7) = arTemp(7)
					arCMT(j,8) = arTemp(8)
					arCMT(j,9) = arTemp(9)

				end if
			next					
		next
	case 2
		For i = 0 to  length
			for j = i to length
				if arCMT(i,5) < arCMT(j,5) then
					arTemp(0) = arCMT(i,0)
					arTemp(1) = arCMT(i,1)
					arTemp(2) = arCMT(i,2)
					arTemp(3) = arCMT(i,3)
					arTemp(4) = arCMT(i,4)
					arTemp(5) = arCMT(i,5)
					arTemp(6) = arCMT(i,6)
					arTemp(7) = arCMT(i,7)
					arTemp(8) = arCMT(i,8)
					arTemp(9) = arCMT(i,9)
						
					arCMT(i,0) = arCMT(j,0)
					arCMT(i,1) = arCMT(j,1)
					arCMT(i,2) = arCMT(j,2)
					arCMT(i,3) = arCMT(j,3)
					arCMT(i,4) = arCMT(j,4)
					arCMT(i,5) = arCMT(j,5)
					arCMT(i,6) = arCMT(j,6)
					arCMT(i,7) = arCMT(j,7)
					arCMT(i,8) = arCMT(j,8)
					arCMT(i,9) = arCMT(j,9)

					arCMT(j,0) = arTemp(0)
					arCMT(j,1) = arTemp(1)
					arCMT(j,2) = arTemp(2)
					arCMT(j,3) = arTemp(3)
					arCMT(j,4) = arTemp(4)
					arCMT(j,5) = arTemp(5)
					arCMT(j,6) = arTemp(6)
					arCMT(j,7) = arTemp(7)
					arCMT(j,8) = arTemp(8)
					arCMT(j,9) = arTemp(9)
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
  %>
  <tr <%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%> >
    <td align="center" class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>"><%=i%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<b><%=arCMT(i,1)%></b><br>
	<i>CMND</i>:<%=arCMT(i,0)%></td>
    <td valign="middle" style="<%=setStyleBorder(0,1,0,1)%>"><%=arCMT(i,2)%></td>
	<td style="<%=setStyleBorder(0,1,0,1)%>" class="CTxtContent"><%=arCMT(i,3)%></td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=day(arCMT(i,4))%>/<%=Month(arCMT(i,4))%>/<%=Year(arCMT(i,4))%></td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=arCMT(i,5)%></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(arCMT(i,6))%></td>
   	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		if 	arCMT(i,7) > 0 then
			Response.Write("Mới: "&arCMT(i,7)&"<br>")
		end if
		
		if arCMT(i,8) > 0 then
			Response.Write("Đang xử lý: "&arCMT(i,8)&"<br>")		
		end if
		
		if arCMT(i,9) > 0 then 
			Response.Write("Hủy: "&arCMT(i,9)&"<br>")		
		end if
	%>
	</td>
   	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><a href="javascript: winpopup('History.asp','<%=arCMT(i,0)%>&Name=<%=arCMT(i,1)%>',990,600);">Chi tiết</a></td>
  </tr>
  <%
	Next
End if

%>
</table>
<%
END IF
%>

<%Call Footer()%>

</body>
</html>
