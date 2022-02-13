<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
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
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	
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
	Title_This_Page="Khách hàng -> Chăm sóc khách hàng"
	img ="../../images/icons/icon_customer1.gif"
	Call header()
	Call Menu()
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	    <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fKhachHang">	
	  <table width="95%" align="center" cellpadding="2" cellspacing="2">
        <tr>
          <td align="right" valign="middle" ><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ngày chăm sóc:</strong></font>
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
          <td align="center" valign="middle" class="CTxtContent" ><em>Điều kiện: </em>
              <input name="txtDieuKien" type="text" id="txtDieuKien" value="<%=strDieuKien%>">
              <em>Tìm theo:</em>
              <select name="selDieuKien" id="selDieuKien">
                <option value="0" selected <%if iDieuKien = 0 then%>selected<%end if%>></option>
                <option value="1" <%if iDieuKien = 1 then%>selected<%end if%>>Tên khách</option>
                <option value="2" <%if iDieuKien = 2 then%>selected<%end if%>>Email</option>
                <option value="3" <%if iDieuKien = 3 then%>selected<%end if%>>Điện thoại</option>
				<option value="3" <%if iDieuKien = 4 then%>selected<%end if%>>Địa chỉ</option>
              </select>          </td>
        </tr>
      
        <tr>
          <td align="center" valign="middle" width="100%" >
              <input name="cbAll" type="checkbox" id="cbAll" value="1">
            Tìm tất cả 
            <input type="submit" name="ButtonSearch" id="ButtonSearch" value="Thống kê">
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
<%
IF Request.form("action")="Search"  THEN
%>
<br><br>

<table width="998" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent" style="border:#CCCCCC solid 1px;">
          <tr>
            <td width="100%" align="center" class="CTieuDe"><p><img src="../../images/icons/icon_customer_service.gif" width="50" height="50" align="absmiddle">DANH SÁCH KHÁCH HÀNG ĐÃ CHĂM SÓC</p></td>
          </tr>
		  
		  <%
		sql = "SELECT CustomerCare.*,Email.ID, Email.Ten, Email.NgaySinh, Email.Diachi, Email.Email FROM CustomerCare INNER JOIN Email ON CustomerCare.IDEmail = Email.ID " 
		sql=sql+" where (DATEDIFF(dd,DateCare,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,DateCare,'" & ToDate &"') >= 0) "
	select case iDieuKien 
		case 1
			sql = sql + " and ({fn UCASE(Email.Ten)} like N'%"& UCase(strDieuKien) &"%') "
		case 2
			sql = sql + " and ({fn UCASE(Email.Email)} like N'%"& UCase(strDieuKien) &"%') "
		case 3
			sql = sql + " and ({fn UCASE(Email.DienThoai)} like N'%"& UCase(strDieuKien) &"%') "
		case 4
			sql = sql + " and ({fn UCASE(Email.Diachi)} like N'%"& UCase(strDieuKien) &"%') "		

	end select		
		sql = sql+" ORDER BY CustomerCare.IDCare DESC"	
		'Response.Write(sql)		
			Set rsCare= Server.CreateObject("ADODB.Recordset")
			rsCare.open sql,Con,3
%>			
          <tr>
            <td valign="middle" background="../../images/TabChinh.gif" style="background-repeat:no-repeat;" height="29" class="CTieuDeNho">&nbsp;Chăm sóc khách hàng</td>
          </tr>
          <tr>
            <td valign="middle" align="center">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" class="CTxtContent" align="center">
			  <tr bgcolor="#FFFF99"> 
				<td width="5%" align="center" class="CTieuDeNhoNho" height="30" style="<%=setStyleBorder(1,1,1,1)%>">Ngày</td>
				<td width="11%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tên KH </td>
				<td width="24%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Địa chỉ </td>
				<td width="22%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Đánh giá</td>
				<td width="5%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Lịch sử </td>
				<td width="10%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">CSKH</td>
				<td width="6%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Thưởng</td>
			    <td width="17%" align="center" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Lý do thưởng </td>
		      </tr>

			<%
				sTT=0
				do while not rsCare.eof
					IDCare	=	rsCare("IDCare")
					DateCare=rsCare("DateCare")
					IDNhanVien=rsCare("IDNhanVien")
					Ghichu=rsCare("Ghichu")
					TThuong=rsCare("TThuong")
					LyDo	=	rsCare("Lydo")
					Ten		=	rsCare("Ten")
					Ngaysinh=	rsCare("NgaySinh")
					Diachi	=	rsCare("Diachi")
					Email	=	rsCare("Email")
					
					Set rsTemp = Server.CreateObject("ADODB.Recordset")
					sqlTemp="SELECT   Account.CMND FROM Account INNER JOIN Email ON Account.Email = Email.Email INNER JOIN CustomerCare ON Email.ID = CustomerCare.IDEmail WHERE IDCare='" & IDCare &"'"
					rsTemp.open sqlTemp,con,1
					iAccount=false
					if  not rsTemp.eof then
					iAccount=true
					CMND=rsTemp("CMND")
					end if	
					rsTemp.close
			%>		
			  <tr>
				<td valign="top" style="<%=setStyleBorder(1,1,0,1)%>" align="center">
				<%=Day(DateCare)%>/<%=month(DateCare)%>/<%=year(DateCare)%></td>
				<td valign="top" style="<%=setStyleBorder(0,1,0,1)%>">
				<%=Ten%><br>
				<font class="CSubTitle">
				NS <%=day(Ngaysinh)%>/<%=month(Ngaysinh)%>/<%=Year(Ngaysinh)%></font>
				</td>
				<td valign="top" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
	iCat	=	instr(Email,"@yahoo.com")
	if iCat > 0 then
			nicname	=	Left(Email,iCat-1)
	%>
	<a title="Email: <%=Email%>" href="ymsgr:sendIM?<%=nicname%>" class="CSubMenu">
	<img src=http://opi.yahoo.com/online?u=<%=nicname%>&m=g&t=2 width=125 height=25 border=0  alt="<%=stralt%>"/></a>
	<%else%>
	<a title="Email: <%=Email%>" href="send_mail.asp?email=<%=Email%>" class="CSubMenu">
	<%
		lg = len(Email)
		temp=""
		if lg > 25 then
			temp	=	Left(Email,15)
			temp	=	temp+".."
		else
			temp =Email
		end if
		Response.Write(temp)
			
			
	%>
	</a>
	<%end if%><br>
	<%if Email <> "" then%>
	<div class="CSubTitle" align="right">
	<a title="Email: <%=Email%>" href="send_mail.asp?email=<%=Email%>" class="CSubMenu">
	Gửi email
	</a>
	</div>
	<%end if%>	
				<font class="CSubTitle"><%=Diachi%></font>				</td>
				<td valign="top" style="<%=setStyleBorder(0,1,0,1)%>">
				<%=GhiChu%>				</td>
				<td  valign="top" style="<%=setStyleBorder(0,1,0,1)%>"><a href="javascript: winpopup('History_email.asp','<%=Email%>&Name=<%=Ten%>',990,600);" class="CSubMenu">Lịch sử</a></td>
				<td valign="top" style="<%=setStyleBorder(0,1,0,1)%>"><%=Response.Write(GetNameNV(IDNhanVien))%> </td>
				<td align="right" valign="top" style="<%=setStyleBorder(0,1,0,1)%>">
				<%if TThuong <> 0 then%>
				<%=Dis_str_money(TThuong)%>
				<%else%>
					<a href="javascript: winpopup('customer_tk.asp','<%=CMND%>',800,600);">Tài khoản</a>
				<%end if%>				</td>
			    <td valign="top" style="<%=setStyleBorder(0,1,0,1)%>"><%=LyDo%>&nbsp;</td>
		      </tr>
			  <%
			  	sTT=sTT+1
			  	rsCare.movenext
			  loop
			  rsCare.close
			  %>			  
			</table>			</td>
          </tr>
</table>
<%end if%>
</body>
</html>
