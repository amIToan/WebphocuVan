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
	img	="../../images/icons/icon_customer1.gif"
	Title_This_Page="Khách hàng -> Thông tin khách hàng"
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
          <td align="center" valign="middle" class="CTxtContent" >Sắp xếp: 
            <select name="selSapXep" id="selSapXep">
              <option value="0" <%if iSapXep = 0 then%>selected<%end if%>></option>
              <option value="1" <%if iSapXep = 1 then%>selected<%end if%>>Họ và tên</option>
              <option value="2" <%if iSapXep = 2 then%>selected<%end if%>>Lượt truy cập</option>
              <option value="3" <%if iSapXep = 3 then%>selected<%end if%>>Ngày truy cập</option>
              <option value="4" <%if iSapXep = 4 then%>selected<%end if%>>Ngày tạo</option>
                                    </select>
            <input name="raTangPricem" type="radio" value="1" checked <%if iTangPricem = 1 then Response.Write("checked") end if %>>
            Giảm dần
            <input name="raTangPricem" type="radio" value="2" <%if iTangPricem = 2 then Response.Write("checked") end if %>> 
            Tăng dần
</td>
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

<%
IF Request.form("action")="Search"  THEN
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)

	
	Dim rs
	Dim emailx
	set rs=server.CreateObject("ADODB.Recordset")
	sql="select top 15 * from Account"
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
			case 0
				sql=sql & " order by CreationDate"			
			case 1
				sql=sql & " order by Name"
			case 2
				sql=sql & " order by iTruyCap"
			case 3
				sql=sql & " order by LastLoginDate"
			case 4		
				sql=sql & " order by CreationDate"			
		end select
		
		if iTangPricem = 1 then
			sql = sql & " DESC "
		end if

	
	rs.open sql,con,1
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</font></strong></td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	else
	i=0
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr align="center" bgcolor="FFFFFF">
    <td width="36" style="<%=setStyleBorder(1,1,1,1)%>"><strong>TT</strong></td>
    <td width="216" style="<%=setStyleBorder(0,1,1,1)%>"><strong>CMND</strong></td>
    <td width="565" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Email</strong></td>
    <td width="118" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tel</strong></td>
    <td width="41" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Số lần </strong></td>
    <td width="87" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Truy cập </strong></td>
    <td width="93" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày tạo</strong></td>
    <td width="95">&nbsp;</td>
  </tr>
  
  <%
  stt = 1
  Do while not rs.eof
  %>
  <tr <%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%> >
    <td align="right" class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<b><%=rs("Name")%></b><br>
	<i>CMND</i>:<%=rs("CMND")%></td>
    <td valign="middle" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if Session("user")="adm" or Session("user")="paparima" then%>
		<a href="send_mail.asp?email=<%=rs("Email")%>">	<%=rs("Email")%></a>
	<%
		if emailx<>"" then
			emailx=emailx+","+rs("Email")
		else
			emailx=rs("Email")
		end if
	else 
		Response.Write(rs("Email"))
	end if
	%>
	<br>
	<font class="CSubTitle">
	<%=rs("diachi")%></font>	</td>
	<td style="<%=setStyleBorder(0,1,0,1)%>" class="CTxtContent">
	<%=rs("Tell")%><br>
	<%=rs("mobile")%>	</td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("iTruyCap")%></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
      <%if not IsNull(rs("LastLoginDate")) then%>
      <%=Hour(ConvertTime(rs("LastLoginDate")))%>:<%=Minute(ConvertTime(rs("LastLoginDate")))%>' <%=Day(ConvertTime(rs("LastLoginDate")))%>/<%=Month(ConvertTime(rs("LastLoginDate")))%>/<%=Year(ConvertTime(rs("LastLoginDate")))%>
      <%Else%>
      &nbsp;
      <%End if%>	</td>
    	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%> </td>
    	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
			<a href="javascript: winpopup('customer_tk.asp','<%=rs("CMND")%>',800,600);">Tài khoản</a><br> 
			<a href="javascript: winpopup('Customer_edit.asp','<%=rs("CMND")%>',800,600);">Sửa</a>|
			 <%if Session("iQuanTri") = 1 then %>
			<a href="javascript: winpopup('Customer_delete.asp','<%=rs("CMND")%>',200,300);">Xóa</a>
			<%end if%>
			 </font>		</td>
  </tr>
  <%i=i+1
  stt=stt+1
  rs.movenext
  Loop
End if
rs.close
set rs=nothing%>
</table>
<%if Session("user")="adm" then%>
		<a href="send_mail.asp?email=<%=emailx%>">(Send all email)</a>
<%
	end if

END IF	
%>
<%Call Footer()%>

</body>
</html>
