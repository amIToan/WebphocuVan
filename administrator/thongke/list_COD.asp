<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_cod")
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

	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)

	iTim 		=	GetNumeric(Request.form("cbAll"),0)
	rdate		=	GetNumeric(Request.form("rdate"),0)
	iQuyetToan	=	GetNumeric(Request.Form("iQuyetToan"),0)
	NVThutienID		=	GetNumeric(Request.form("selNVThutien"),0)
	DoiTuongThuTien	=	getNhanVienFromID(NVThutienID)
	NhanVienID		=	GetNumeric(Request.Form("NhanVienID"),0)
	reportDoiTac	= 	GetNumeric(Request.form("reportDoiTac"),0)
	iMaorTenSach	=	Clng(Request.Form("selMaorTenSach"))
	strMaorTenSach	=	Trim(Request.Form("txtMaOrTensach"))
	StatusDonhang	=	GetNumeric(Request.form("StatusDonhang"),0)
	inotpay			= 	GetNumeric(Request.form("notpay"),0)
	inotpay2			= 	GetNumeric(Request.form("notpay2"),0)
	isort			= 	GetNumeric(Request.form("sel_sort"),0)
	iDownUp			= 	GetNumeric(Request.form("sel_down_up"),0)
ELSE
	Day1 = now() - 30
	Ngay1=Day(Day1)
	Thang1=Month(Day1)
	Nam1=Year(Day1)
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
END IF
%>

<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<SCRIPT language=JavaScript1.2 src="../administrator/inc/calendarDateInput.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Request.form("action")<>"Search"  THEN
	Title_This_Page="Th???ng k??->Danh s??ch COD"
	Call header()
	Call Menu()
	
	
%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td  background="../../images/T1.jpg" height="20"></td>
    </tr>
    <tr>
      <td background="../../images/t2.jpg">
	  <FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fThongke" onSubmit="return checkme();">
  
  <table width="99%" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
    <tr>
      <td align="right" valign="middle" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
        <tr>
          <td width="21%" align="right"><div align="left">Th???i gian t??? ng??y: </div></td>
          <td width="34%">
		<%
			Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call List_Month_WithName(Thang1,"MM","Thang1")
			Call  List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
		%>		  </td>
          <td width="11%" align="right">?????n ng??y:</td>
          <td width="34%">
		<%
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call List_Month_WithName(Thang2,"MM","Thang2")
			Call  List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
		%>		  </td>
        </tr>

      </table></td>
    </tr>
        <tr>
          <td >Th???ng k??? theo: 
            <input name="rdate" type="radio" value="0" checked>
            Ng??y x??? l??  &nbsp;&nbsp;&nbsp; 
            <input name="rdate" type="radio" value="1"> 
            Ng??y thanh to??n</td>
          </tr>	
    <tr> 
   <tr>
      <td valign="middle" >
	    ??i???u ki???n:
	        <input name="txtMaOrTensach" type="text" id="txtMaOrTensach" value="<%=strMaorTenSach%>">  
	        T??m theo:	
	        <select name="selMaorTenSach" id="selMaorTenSach">
	          <option value="0" selected <%if iMaorTenSach = 0 then%>selected<%end if%>></option>
	          <option value="1" <%if iMaorTenSach = 1 then%>selected<%end if%>>M?? ????n h??ng</option>
	          <option value="3" <%if iMaorTenSach = 3 then%>selected<%end if%>>T??n kh??ch</option>
	          <option value="4" <%if iMaorTenSach = 4 then%>selected<%end if%>>Email</option>
	          <option value="5" <%if iMaorTenSach = 5 then%>selected<%end if%>>Tel</option>
	          <option value="7" <%if iMaorTenSach = 7 then%>selected<%end if%>>?????a ch???</option>
	            </select>        </td>
    </tr>
    <tr>
      <td valign="middle" style="border-bottom:#99CCFF solid 1px;">S???p x???p:
        <select name="sel_sort" id="sel_sort">
		<option value="0" selected>Ng??y thanh to??n</option>
          <option value="1" selected>Ng??y x??? l??</option>
          <option value="2">S??? Bill</option>
          <option value="3">T??n Kh??ch</option>
          <option value="4">Ng??y ?????t h??ng</option>
        </select>
         <select name="sel_down_up" id="sel_down_up">
           <option value="0">T??ng d???n</option>
           <option value="1">Gi???m d???n</option>
         </select>      </td>
    </tr>			
    <tr>
      <td  valign="middle" class="CTxtContent" >
	    Giao h??ng:
	      <%
			call SelectNhanVien("NhanVienID",NhanVienID,6,0,0)
			%>&nbsp;&nbsp;&nbsp;  Thu ti???n:
	      <%
			call SelectNhanVien("selNVThutien",NVThutienID,6,0,0)
			%>
		</td>
    </tr>	
    <tr>
      <td valign="middle" >
          <input name="notpay" type="checkbox" value="1">
          Ch??a thanh to??n&nbsp;&nbsp;&nbsp;
          <input name="notpay2" type="checkbox" value="1">
          ???? thanh to??n  &nbsp;&nbsp;&nbsp;
  <input name="reportDoiTac" type="checkbox" id="reportDoiTac" value="1">
	      TK d??nh cho ?????i t??c&nbsp;&nbsp;&nbsp;          <input name="iQuyetToan" type="checkbox" id="iQuyetToan" value="1">
          Quy???t to??n COD </td>
    </tr>

    <tr>
      <td align="center" valign="middle" >
	   <input name="cbAll" type="checkbox" id="cbAll" value="1">
          T??m t???t c???         <%
			call ListStatusOfDonhang(StatusDonhang)
		%>
          <input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
          <input type="hidden" name="action" value="Search">
          <input type="hidden" name="OrderType" value=""></td>
    </tr>
  </table>
</form>
	  </td>
    </tr>
    <tr>
      <td background="../../images/T3.jpg" height="8"></td>
    </tr>
</table>
<br> <center><img src="../../images/line5.gif" height="1" ><img src="../../images/line5.gif" height="1" ></center><br>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function order(OrderType)
 {
 	if (!checkme())
 		return;
 	document.fThongke.OrderType.value=OrderType;
 	document.fThongke.submit();
 }
 function checkme()
 {
	if (document.fThongke.StatusDonhang.value==-1)
	{
		alert("B???n h??y ch???n lo???i ????n h??ng!");
		document.fThongke.StatusDonhang.focus();
		return false;
	}
	return true;
 }
// -->
</SCRIPT>

<%
end if
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
IF Request.form("action")="Search"  THEN
	Dim rstemp
	Set rstemp=Server.CreateObject("ADODB.Recordset")
	sql="SELECT count(SanPhamUser_ID) as iCount FROM V_SanPham_Donhang where SanPhamUser_Status="&StatusDonhang&"  "
	rstemp.open sql,con,1
	if not rstemp.eof then
		STT = rstemp("iCount")
	else
		STT = 0
	end if

	Dim rs

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM V_SanPham_Donhang where " 
	if iTim = 0  then
		sql=sql+" SanPhamUser_Status="&StatusDonhang&" and "
	end if
	if NhanVienID <> 0 then
		sql=sql+" NhanVienID = "& NhanVienID&" and "
	end if
	if NVThutienID <> 0 then
		sql=sql+" NVThutienID = "& NVThutienID &" and "
	end if	
	if rdate = 1 then
		sql=sql+" (DATEDIFF(dd,NgayThanhToan,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayThanhToan,'" & ToDate &"') >= 0) "
	else
		sql=sql+" (DATEDIFF(dd,NgayXuLy,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayXuLy,'" & ToDate &"') >= 0) "
	end if
	
	if inotpay =  1 and inotpay2 <> 1 then 
		sql =sql+" and (DATEDIFF(dd,NgayThanhToan,'" & FormatDatetime("1/1/2005") & "')> 0) and NgayThanhToan='' "
	end if
	if inotpay2 =  1 and inotpay <>  1 then 
		sql =sql+" and (DATEDIFF(dd,NgayThanhToan,'" & FormatDatetime("1/1/2005") & "')<= 0) and NgayThanhToan<>'' "
	end if	
	select case iMaorTenSach 
		case 1
			strMaorTenSach = 	replace(strMaorTenSach,"XB","")			
			if isnumeric(strMaorTenSach) = true then
				numb = Clng(strMaorTenSach) - 1000
			else
				numb = 0
			end if	
			sql = sql + " and SanPhamUser_ID = "&numb
		case 3
			sql = sql + " and {fn UCASE(SanPhamUser_Name)} like N'%"& UCase(strMaorTenSach) &"%'"
		case 4
			sql = sql + " and SanPhamUser_Email like N'%"& strMaorTenSach &"%'"
		case 5
			sql = sql + " and SanPhamUser_Tell like N'%"& strMaorTenSach &"%'"			
		case 7
			sql = sql + " and {fn UCASE(GiaoHang_Address)} like N'%"& UCase(strMaorTenSach) &"%'"			
			
	end select		

	strTemp	=	"NgayXuLy"
	select case isort
		case 0 
			strTemp	= "NgayThanhToan"
		case 1
			strTemp = "NgayXuLy"
		case 2
			strTemp = "MaBuuDien"
		case 3
			strTemp = "SanPhamUser_Name"
		case 4
			strTemp = "SanPhamUser_Date"	
	end select
	if iDownUp =  1 then
		strDesc =	" DESC"	
	end if
	
	sql=sql+"ORDER BY " + strTemp + strDes
	rs.open sql,con,3
	call UserOperation(user,hour(now)&":"&Minute(now)&"phut : th???ng k?? COD "&rs.recordcount)
	if rs.eof then 'Kh??ng c?? b???n ghi n??o th???a m??n
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</font></strong></td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	else

%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
    <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
    <em>??T: <%=soDT%> - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>??C: <%=dcVanPhong%> </em></div></td>
  </tr>
</table>
<br><br>
  <div align="center"class="author">
    <div align="center"><strong>THANH TO??N COD </strong></div>
  </div>
  <center> T??? ng??y <%=Day(FromDate)%>/<%=month(FromDate)%>/<%=Year(FromDate)%> ?????n <%=Day(ToDate)%>/<%=month(ToDate)%>/<%=Year(ToDate)%></center>
<br>
 <div align="center" class="author">
  <div align="left"><strong><font class="CTxtContent">????n v???: </font><%=DoiTuongThuTien%></strong> </div>
</div> 
<br>
<%if iQuyetToan =  0 then%>
<table  border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
  <tr bgcolor="#CCFFFF">
    <td align="center" style="<%=setStyleBorder(1,1,1,1)%>"><strong>Ng??y g???i </strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>S???</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>M??</strong></td>
<%if reportDoiTac <> 1 then%>	
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>T??n/?????a ch???</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" ><strong>Ki???m so??t</strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>TL Web  </strong></td>
<%end if%>	
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" width="20"><strong>TL Th???c </strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" width="20"><strong> C?????c chi </strong></td>
   <%if reportDoiTac <> 1 then%>
   <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>C?????c Thu </strong></td>
   <%end if%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ti???n thu </strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" width="150"><strong>Ng?????i nh???n</strong></td>
  </tr>
  <%
iMau=0
SoDH = 0
nTongTien =0
nTongCuoc = 0
nTongTT	=	0
t_trong_luong	=	0
t_cuoc_thu	=	0
t_cuoc_chi	=	0
t_trong_luong_thuc=0
Do while not rs.eof 
	SanPhamUser_ID		=	rs("SanPhamUser_ID")
	SanPhamUser_Name	=	rs("SanPhamUser_Name")
	SanPhamUser_Email	=	rs("SanPhamUser_Email")
	SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
	SanPhamUser_Address	=	rs("SanPhamUser_Address")
	SanPhamUser_Thoigian=	rs("SanPhamUser_Thoigian")
	SanPhamUser_Status	=	rs("SanPhamUser_Status")
	NgayXuLy			=	rs("NgayXuLy")
	NgayTT				=	rs("NgayThanhToan")
	strCMND				=	rs("CMND")
	GiaoHang_Address	=	rs("GiaoHang_Address")
	MaBuuDien			=	rs("MaBuuDien")	
	testpayid				=	getNhanVienFromID(rs("testpayid"))
	checkpayid				=	getNhanVienFromID(rs("checkpayid"))
%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%> </td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
			<%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a>	</td>
    <td  valign="middle" align="left" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		if MaBuuDien = "" or MaBuuDien = NULL then
			Response.Write("&nbsp;")
		else
			Response.Write(MaBuuDien)
		end if
	%>	</td>
  <%if reportDoiTac <> 1 then%>
    <td  valign="middle" align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=SanPhamUser_Name%><br>
        <font class="CSubTitle"> <i>?????a ch???</i>: <%=SanPhamUser_Address%><br>
      </font></td>

    <td  align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=testpayid%></td>
    <td  style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		t_trong_luong	=	t_trong_luong+TongTrongLuong(SanPhamUser_ID)
	%>
	<%=TongTrongLuong(SanPhamUser_ID)%>g	</td>
<%end if%>	
    <td  style="<%=setStyleBorder(0,1,0,1)%>">
		<%
		n_trong_luong_thuc	=	GetKhoiLuongThucID(SanPhamUser_ID)
		t_trong_luong_thuc	=	t_trong_luong_thuc+n_trong_luong_thuc
	%>
	<%=n_trong_luong_thuc%>g</td>
    <td width="20" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		nTongCuoc = nTongCuoc + GetCuocBuuDienThucID(SanPhamUser_ID)
		Response.Write(Dis_str_money(GetCuocBuuDienThucID(SanPhamUser_ID)))
	%>	</td>
	<%if reportDoiTac <> 1 then%>
    <td  align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		n_cuoc_thu	=	GetCuocBuuDienID(SanPhamUser_ID)+GetPhiVanChuyen(SanPhamUser_ID)
		t_cuoc_thu	=	t_cuoc_thu + n_cuoc_thu
	%>
	<%=Dis_str_money(n_cuoc_thu)%>	</td>
	<%end if%>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		iTien = 0
		if rs("NVThutienID") = NVThutienID then
			iTien = LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND))
		end if
		Response.Write(Dis_str_money(iTien))
		nTongTien =	nTongTien + iTien
		

	%>	</td>
    <td   style="<%=setStyleBorder(0,1,0,1)%>">
	<%=SanPhamUser_Name%>
	-
	<%
		iPos	=	InStrRev(SanPhamUser_Address,";")+1
		SanPhamUser_Address	=	Mid(SanPhamUser_Address,iPos,Len(SanPhamUser_Address))
		if Len(SanPhamUser_Address)> 12 then
			iPos	=	InStrRev(SanPhamUser_Address,",")+1
			SanPhamUser_Address	=	Mid(SanPhamUser_Address,iPos,Len(SanPhamUser_Address))		
		end if
		Response.Write(TRIM(UCASE(SanPhamUser_Address)))
	%>	</td>
  </tr>

  <%
	SoDH = SoDH+1
	stt=stt - 1
	iMau=iMau+1
	rs.movenext
Loop%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td colspan="5" align="center" style="<%=setStyleBorder(1,1,0,1)%>">T???ng:</td>
    <td  style="<%=setStyleBorder(0,1,0,1)%>"><%=t_trong_luong%></td>
    <td  style="<%=setStyleBorder(0,1,0,1)%>"><%=t_trong_luong_thuc%></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(nTongCuoc)%></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(t_cuoc_thu)%></td>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td   style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
  </tr>
</table>
<%	
	rs.close
	set rs=nothing
%>	
<br> <br>
<%else%>
<form action="updateCOD.asp" target="_blank" method="post" name="ThanhToanCOD">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#CCFFFF">
    <td align="center" style="<%=setStyleBorder(1,1,1,1)%>">M??</td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ng??y g???i </strong></td>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ng??y TT</strong> </td>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ch???n TT </strong></td>
	<%end if%>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>H???y TT</strong> </td>
	<%end if%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>T??n/?????a ch???</strong></td>
	<%if reportDoiTac <> 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" >Ti???p nh???n phi???u</td>    
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>" ><strong>Ki???m so??t TT </strong></td>
    <%end if%>	
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong> C?????c chi </strong></td>
   <%if reportDoiTac <> 1 then%>
   <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>C?????c Thu </strong></td>
   <%end if%>
    <td align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ti???n h??ng </strong></td>
  </tr>
  <%
iMau=0
iSoChonTT=0
iSoHuyTT=	0
SoDH = 0
nTongTien =0
nTongCuoc = 0
nTongTT	=	0
Do while not rs.eof 
	SanPhamUser_ID		=	rs("SanPhamUser_ID")
	SanPhamUser_Name	=	rs("SanPhamUser_Name")
	SanPhamUser_Email	=	rs("SanPhamUser_Email")
	SanPhamUser_Tell	=	rs("SanPhamUser_Tell")
	SanPhamUser_Address	=	rs("SanPhamUser_Address")
	SanPhamUser_Thoigian=	rs("SanPhamUser_Thoigian")
	SanPhamUser_Status	=	rs("SanPhamUser_Status")
	NgayXuLy			=	rs("NgayXuLy")
	NgayTT				=	rs("NgayThanhToan")
	strCMND				=	rs("CMND")
	GiaoHang_Address	=	rs("GiaoHang_Address")
	testpayid				=	getNhanVienFromID(rs("testpayid"))
	checkpayid				=	getNhanVienFromID(rs("checkpayid"))
%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">	<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
			<%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%></a></td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(ConvertTime(NgayXuLy))%>/<%=Month(ConvertTime(NgayXuLy))%>/<%=Year(ConvertTime(NgayXuLy))%> </td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if isdate(NgayTT) = true and NgayTT <> #1/1/1900# then%>
	<%=Day(NgayTT)%>/<%=Month(NgayTT)%>/<%=Year(NgayTT)%>
	<%
		iTien=0	
		if rs("NVThutienID") = NVThutienID then
			iTien = LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND)) - GetCuocBuuDienThucID(SanPhamUser_ID)
		elseif rs("NhanVienID") = NVThutienID then
			iTien =	iTien - GetCuocBuuDienThucID(SanPhamUser_ID)
		end if
		nTongTT =	nTongTT + iTien
		%>
	<%else%>
	<img src="../images/icon-banner-new.gif" height="16" width="16" border="0">
	<%end if%>	</td>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if isdate(NgayTT) <> true or NgayTT = #1/1/1900# then%>
	<input type="checkbox" name="ChonTT<%=iSoChonTT%>" value="1">
	<input type="hidden" name="User_ID_chon<%=iSoChonTT%>" value="<%=SanPhamUser_ID%>">
	
	<%
		iSoChonTT	=	iSoChonTT + 1
	else%>	
		&nbsp;
	<%end if%>	  </td>
	  <%end if%>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if isdate(NgayTT) = true and NgayTT <> #1/1/1900# then%>
	<input type="checkbox" name="HuyTT<%=iSoHuyTT%>" value="1">
	<input type="hidden" name="User_ID_Huy<%=iSoHuyTT%>" value="<%=SanPhamUser_ID%>">
	<%
	iSoHuyTT	=	iSoHuyTT + 1
	else%>	
		&nbsp;
	<%end if%>	</td>
	<%end if%>
    <td  valign="middle" align="left" style="<%=setStyleBorder(0,1,0,1)%>">
	<a href="ReportXKChiTiet.asp?SanPhamUser_ID=<%=SanPhamUser_ID%>" target="_parent" class="CSubMenu">
	<%=SanPhamUser_Name%></a><br>
        <font class="CSubTitle">
		<u>?????a ch???:</u><%=SanPhamUser_Address%><br>
    </font>	</td>
	<%if reportDoiTac <> 1 then%>
    <td  align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=checkpayid%></td>
    
    <td align="left" style="<%=setStyleBorder(0,1,0,1)%>"><%=testpayid%></td>
    <%end if%>	
    <td  align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		nTongCuoc = nTongCuoc + GetCuocBuuDienThucID(SanPhamUser_ID)
		Response.Write(Dis_str_money(GetCuocBuuDienThucID(SanPhamUser_ID)))
	%>	</td>
	<%if reportDoiTac <> 1 then%>
    <td width="9%" align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%=Dis_str_money(GetCuocBuuDienID(SanPhamUser_ID)+GetPhiVanChuyen(SanPhamUser_ID))%>	</td>
	<%end if%>
    <td  align="right" style="<%=setStyleBorder(0,1,0,1)%>">
	<%
		iTien = 0
		if rs("NVThutienID") = NVThutienID then
			iTien = LamTronTien(TongTienTrenDonHang(SanPhamUser_ID,strCMND))
		end if
		Response.Write(Dis_str_money(iTien))
		nTongTien =	nTongTien + iTien
	%>	</td>
  </tr>

  <%
	SoDH = SoDH+1
	stt=stt - 1
	iMau=iMau+1
	rs.movenext
Loop%>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<td align="right" style="<%=setStyleBorder(0,1,0,1)%>">Ch???n t???t: </td>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if iSoChonTT  > 0 then%>
	<input type="checkbox" name="ChonTTAll" value="1" onClick="javascript:OnCheckAll();">
	<input type="hidden" name="iSoTTAll" value="<%=iSoChonTT-1%>">	
	<%else%>
		&nbsp;
	<%end if%>	</td>
	<%end if%>
	<%if f_permission > 1 then%>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">
	<%if iSoHuyTT  > 0 then%>
	<input name="ChonHuyTTAll" type="checkbox" id="ChonHuyTTAll" value="1" onClick="javascript:OnCheckAllHuy();">
	<input type="hidden" name="iSoHuyTTAll" value="<%=iSoHuyTT-1%>">
		<%else%>
		&nbsp;
	<%end if%>	</td>
	<%end if%>
    <td valign="middle" align="left" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%if reportDoiTac <> 1 then%>
    <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    <td align="left" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%end if%>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%if reportDoiTac <> 1 then%>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
	<%end if%>
    <td align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
  </tr>
  <tr <%if iMau mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
    <td colspan="11" align="center" style="<%=setStyleBorder(1,1,0,1)%>">
	Ch???n ng??y c???p nh???t thanh to??n:		<%
			Call List_Date_WithName(Day(Now),"DD","NgayCOD")
			Call List_Month_WithName(month(now),"MM","ThangCOD")
			Call  List_Year_WithName(year(now),"YYYY",2004,"NamCOD")
		%>
	<input type="submit" name="Submit" value="C???p nh???t">
      <input type="reset" name="Submit2" value="L??m l???i">      </td>
    </tr>
</table>
</form>
<br>
<%end if%>
<table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td  background="../../images/T1.jpg" height="20"></td>
  </tr>
  <tr>
    <td background="../../images/t2.jpg">
	  <table width="95%" border="0"  align="center" cellpadding="2" cellspacing="2"  id="table91">
        <tr>
          <td colspan="3" align="center" class="style14"><strong>TH???NG K?? COD </strong><br />
          <img src="../../images/line5.gif" height="1" ></td>
        </tr>
        <tr>
          <td width="49%" align="right" class="CTxtContent">T???ng s??? l?????ng h??ng g???i: </td>
          <td width="23%" align="left" class="CTxtContent"><%=SoDH%></td>
          <td width="28%" align="right" class="CTxtContent">&nbsp;</td>
        </tr>
        <tr>
          <td align="right" class="CTxtContent">T???ng ti???n c?????c: </td>
          <td align="left" class="CTxtContent"><%=Dis_str_money(LamTronTien(nTongCuoc))%></td>
          <td align="right" class="CTxtContent">&nbsp;</td>
        </tr>
        <tr>
          <td align="right" class="CTxtContent">T???ng ti???n h??ng: </td>
          <td align="left" class="CTxtContent" style="<%=setStyleBorder(0,0,0,1)%>"><%=Dis_str_money(LamTronTien(nTongTien))%></td>
          <td align="right" class="CTxtContent">&nbsp;</td>
        </tr>
        <tr>
          <td align="right" class="CTxtContent"><strong> T???ng thanh to??n: </strong></td>
          <td align="left" class="CTxtContent"><strong><%=Dis_str_money(LamTronTien(nTongTien-nTongCuoc))%></strong></td>
          <td align="right" class="CTxtContent">&nbsp;</td>
        </tr>
		<%if iQuyetToan <>  0 then%>
        <tr>
          <td align="right" class="CTxtContent">???? thanh to??n: </td>
          <td align="left" class="CTxtContent" style="<%=setStyleBorder(0,0,0,1)%>"><%=Dis_str_money(LamTronTien(nTongTT))%></td>
          <td align="right" class="CTxtContent">&nbsp;</td>
        </tr>
        <tr>
          <td align="right" class="CTxtContent"><strong><%=DoiTuongThuTien%> Thanh to??n c??n l???i: </strong></td>
          <td align="left" class="CTxtContent"><b><%=Dis_str_money(LamTronTien(nTongTien-nTongCuoc-nTongTT))%></b></td>
          <td align="right" class="CTxtContent">&nbsp;</td>
        </tr>
		<%end if%>
        <tr>
          <td colspan="3" align="center" class="CTxtContent"><strong>Ghi b???ng ch???: <%=tienchu(LamTronTien(nTongTien-nTongCuoc-nTongTT))%></strong></td>
        </tr>
      </table>	</td>
  </tr>
  <tr>
    <td background="../../images/T3.jpg" height="8">
	</td>
  </tr>
</table>
<br>
<%
end if 'if not rs.eof then
END IF 'IF Request.form("action")="Search" THEN

IF Request.form("action")<>"Search"  THEN
Call Footer()
end if
%>
</body>
</html>
<script language="javascript">
function OnCheckAll()
{
	iNumSP	=	document.ThanhToanCOD.iSoTTAll.value;	
	for(jj=0;jj<=iNumSP;jj++)
	{
		if (document.ThanhToanCOD.ChonTTAll.checked == true)
			str = "document.ThanhToanCOD.ChonTT"+jj+".checked = true";
		else
			str = "document.ThanhToanCOD.ChonTT"+jj+".checked = false";
		eval(str);	
	}
	return;
	
}
function OnCheckAllHuy()
{
	iNumSP	=	document.ThanhToanCOD.iSoHuyTTAll.value;	
	for(jj=0;jj<=iNumSP;jj++)
	{
		if (document.ThanhToanCOD.ChonHuyTTAll.checked == true)
			str = "document.ThanhToanCOD.HuyTT"+jj+".checked = true";
		else
			str = "document.ThanhToanCOD.HuyTT"+jj+".checked = false";
		eval(str);	
	}
	return;
	
}
</script>