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
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
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
	iGop=GetNumeric(Request.form("iGop"),0)
	iDonHang=GetNumeric(Request.form("iDonHang"),0)
	StatusDonhang	=	GetNumeric(Request.form("StatusDonhang"),0)
	TinhID		=	GetNumeric(Request.form("selTinh1"),0)
ELSE
	Day1 = now() - 30
	Ngay1=Day(Day1)
	Thang1=Month(Day1)
	Nam1=Year(Day1)
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
	StatusDonhang=Clng(Request.QueryString("StatusDonhang"))
END IF
%>

<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Request.form("action")<>"Search"  THEN
	Title_This_Page="Thống kê -> theo địa lý"
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
  
  <table width="95%" align="center" cellpadding="2" cellspacing="2">
    <tr>
      <td align="right" valign="middle" class="CTxtContent" >
	  	  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
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
      <td align="center" valign="middle" class="CTxtContent" >
	<select name="selTinh1"  onChange="CopyTinh();">
	<option value="0" selected>Tất cả các tỉnh thành</option>
		<%
		sql="Select * From Tinh"
		set rspr=Server.CreateObject("ADODB.Recordset")
		rspr.open sql,Con,1 
		Do while not rspr.eof
			%>
			<option value="<%=rspr("TinhID")%>" <%if TinhID = rspr("TinhID") then Response.Write("selected=""selected""") end if %>> <%=rspr("TenTinh")%></option>
			<%
		rspr.movenext
		Loop
		rspr.close
		%>
	</select></td>
    </tr>
    
    <tr>
      <td  valign="middle" class="CTxtContent"  align="center"><input name="iGop" type="checkbox" id="iGop" value="1">
        Gộp đơn hàng và doanh thu &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Địa danh có 
          <input name="iDonHang" type="text" id="iDonHang" value="0" size="2" maxlength="3" onBlur="javascrip:checkIsNumber(this)">
          đơn hàng trở lên </td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="100%" ><div align="center" class="CTxtContent">
        <input name="cbAll" type="checkbox" id="cbAll" value="1">
        Tìm tất cả / 
		  <%
			call ListStatusOfDonhang(StatusDonhang)
		%>
		    <input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
		    <input type="hidden" name="action" value="Search">
		    <input type="hidden" name="OrderType" value="">	  
		  </div></td>
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
		alert("Bạn hãy chọn loại đơn hàng!");
		document.fThongke.StatusDonhang.focus();
		return false;
	}
	return true;
 }
// -->
</SCRIPT>

<%
end if
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)

IF Request.form("action")="Search"  THEN
%>
	<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
	  <tr>
		<td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
		<td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
		<em>ĐT: <%=soDT%> - Email: info@xseo.com</em></td>
	  </tr>
	  <tr>
		<td><div align="center"><strong><%=TenGD%></strong></div></td>
		<td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%> </em></div></td>
	  </tr>
	</table>
	<br>
	  <div align="center"class="author">
		<div align="center"><strong>THỐNG KÊ THEO ĐỊA LÝ </strong></div>
	  </div>
	  <center> Từ ngày <%=Ngay1%>/<%=Thang1%>/<%=Nam1%> Đến <%=Ngay2%>/<%=Thang2%>/<%=Nam2%></center>
	<BR>
<%	
		if TinhID = 0 then
			sql="Select * From Tinh"
			set rspr=Server.CreateObject("ADODB.Recordset")
			rspr.open sql,Con,1 
			iCount 	=	rspr.recordcount-1
			redim	arTinh(iCount)
			redim	arTinhValue(iCount)
			redim 	arFTinh(iCount)
			h = 0
			Do while not rspr.eof
				arTinh(h) = rspr("TenTinh")
				h=h+1
				rspr.movenext
			Loop
			rspr.close
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql="SELECT SanPhamUser_ID,CMND,SanPhamUser_Address FROM V_SanPham_Donhang where " 
			if iTim = 0  then
			sql=sql+" SanPhamUser_Status="&StatusDonhang&" and "
			end if
			sql=sql+" (DATEDIFF(dd,NgayXuLy,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayXuLy,'" & ToDate &"') >= 0) " 
			sql=sql+"ORDER BY SanPhamUser_ID DESC"
			rs.open sql,con,3
			do while not rs.eof 
				for h = 0 to iCount 
					iCheck=InStr(rs("SanPhamUser_Address"),arTinh(h))
					if  iCheck > 0 then
						arTinhValue(h) = arTinhValue(h)+1
						arFTinh(h)		= arFTinh(h) + TongTienTrenDonHang(rs("SanPhamUser_ID"),rs("CMND"))
					end if 
				next
				rs.movenext
			loop
			set rs = nothing
		else
			sql="Select * From Tinh where TinhID="&TinhID
			set rspr=Server.CreateObject("ADODB.Recordset")
			rspr.open sql,Con,1 
			if not rspr.eof then
				TenTinh	= rspr("TenTinh")
			end if
			set rspr = nothing

			sql = "Select * from Huyen where TinhID="&TinhID
			set rspr=Server.CreateObject("ADODB.Recordset")
			rspr.open sql,Con,1 
			iCount	= rspr.recordcount-1
			redim 	arHuyen(iCount)
			redim 	arHuyenValue(iCount)
			redim 	arFHuyen(iCount)
			h=0
			do while not rspr.eof 
				arHuyen(h)	= rspr("TenHuyen")
				h=h+1
				rspr.movenext	
			loop
			set rspr = nothing
			
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql="SELECT SanPhamUser_ID,CMND,SanPhamUser_Address FROM V_SanPham_Donhang where " 
			if iTim = 0  then
			sql=sql+" SanPhamUser_Status="&StatusDonhang&" and "
			end if
			sql=sql+" (DATEDIFF(dd,NgayXuLy,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgayXuLy,'" & ToDate &"') >= 0) " 
			sql=sql+"ORDER BY SanPhamUser_ID DESC"
			rs.open sql,con,3
			do while not rs.eof 
				for h = 0 to iCount 
					iCheck	= InStr(rs("SanPhamUser_Address"),TenTinh)
					iCheck1	= InStr(rs("SanPhamUser_Address"),arHuyen(h))
					if  iCheck > 0 and iCheck1 > 0  then
						arHuyenValue(h) = arHuyenValue(h)+1
						arFHuyen(h)		= arFHuyen(h) + TongTienTrenDonHang(rs("SanPhamUser_ID"),rs("CMND"))
					end if 
				next
				rs.movenext
			loop
	
		end if


%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td scope="row">&nbsp;</td>
  </tr>
  <tr>
    <td scope="row">
	<%if TinhID = 0 then%>
	<table  border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td>
		<%if iGop = 1 then %>
			<table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
          <tr>
            <td style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1" align="right">&nbsp;</td>
            <td style="border-bottom:#000000 solid 1">&nbsp;</td>
            <td align="right" style="border-bottom:#000000 solid 1">Số ĐH / Triệu </td>
          </tr>
          <%
	   fTongDH = 0
	   	for i = 0 to iCount
			if arTinhValue(i) >= iDonHang  then
			fTongDH	= fTongDH+arTinhValue(i)
			fTong	=	fTong+arFTinh(i)

		%>
          <tr>
            <td style="border-right:#000000 solid 1" align="right"><%=arTinh(i)%>:</td>
            <td height="50"><%for k = 0 to arTinhValue(i)
			Response.Write("<img src=""../../images/TKeNgang.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
              &nbsp; <%=GetNumeric(arTinhValue(i),0)%> <br>
              <%for k = 0 to round(arFTinh(i)/100000)
			Response.Write("<img src=""../../images/TKeNgangDo.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
              &nbsp; <%=round(arFTinh(i)/1000000,3)%> </td>
            <td>&nbsp;</td>
          </tr>
          <%
			end if
		  next%>
          <tr>
            <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Tỉnh</td>
            <td>Tổng đơn hàng: <%=fTongDH%><br>
              Tổng tiền: <%=Dis_str_money(fTong)%> đ </td>
            <td valign="top">
				<img src="../../images/TKeNgang.jpg" width="12" height="12"> Số đơn hàng<br>
				<img src="../../images/TKeNgangDo.jpg" width="12" height="12"> Doanh thu			 </td>
          </tr>
        </table>
          <%else%>
		  
		  <table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
          <tr>
            <td style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1" align="right">Tỉnh\ĐH</td>
            <td style="border-bottom:#000000 solid 1">&nbsp;</td>
            <td align="right" style="border-bottom:#000000 solid 1">Đơn hàng </td>
          </tr>
          <%
	   fTongDH = 0
	   	for i = 0 to iCount
			if arTinhValue(i) >=iDonHang then
			fTongDH	= fTongDH + arTinhValue(i)

		%>
          <tr>
            <td style="border-right:#000000 solid 1" align="right"><%=arTinh(i)%>:</td>
            <td height="25"><%for k = 0 to arTinhValue(i)
			Response.Write("<img src=""../../images/TKeNgang.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
              &nbsp; <%=GetNumeric(arTinhValue(i),0)%> <br>
			  
			  </td>
            <td>&nbsp;</td>
          </tr>
          <%
		  	end if
		  next%>
          <tr>
            <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Tỉnh</td>
            <td>Tổng: <%=fTongDH%></td>
            <td>&nbsp;</td>
          </tr>
        </table>
		
          <table border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
            <tr>
              <td width="25" >&nbsp;</td>
              <td width="99" align="right" style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1">Tỉnh\Doanh thu </td>
              <td width="67" style="border-bottom:#000000 solid 1">&nbsp;</td>
              <td width="54" align="right" style="border-bottom:#000000 solid 1">Triệu</td>
            </tr>
            <%
	   fTong = 0
	   	for i = 0 to iCount
			if arTinhValue(i) >=  iDonHang then
			fTong	= fTong+arFTinh(i)
		%>
            <tr>
              <td >&nbsp;</td>
              <td style="border-right:#000000 solid 1" align="right"><%=arTinh(i)%>:</td>
              <td height="25"><%for k = 0 to round(arFTinh(i)/100000)
			Response.Write("<img src=""../../images/TKeNgangDo.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
                &nbsp; <%=round(arFTinh(i)/1000000,3)%> </td>
              <td>&nbsp;</td>
            </tr>
            <%
				end if
			next%>
            <tr>
              <td>&nbsp;</td>
              <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Tỉnh</td>
              <td>Tổng: <%=Dis_str_money(fTong)%> đ</td>
              <td>&nbsp;</td>
            </tr>
          </table>
		  <%end if%>
		  </td>
      </tr>
    </table>
	<%else%>
	<table  border="0" cellspacing="2" cellpadding="2" align="center">
      <tr>
        <td align="center"><span class="CFontVerdana10">Thống kê trên địa bàn: <%=TenTinh%></span></td>
      </tr>
      <tr>
        <td>
		<%if iGop = 1 then%>
	  <table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
        <tr>
          <td style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1" align="right">&nbsp;</td>
          <td style="border-bottom:#000000 solid 1">&nbsp;</td>
          <td align="right" style="border-bottom:#000000 solid 1">Số ĐH / Triệu</td>
        </tr>
        <%
	   fTongDH = 0
	   	for i = 0 to iCount
			if arHuyenValue(i) >=  iDonHang then
			fTongDH	= fTongDH+arHuyenValue(i)
			fTong	=	fTong+arFHuyen(i)

		%>
        <tr>
          <td style="border-right:#000000 solid 1" align="right"><%=arHuyen(i)%>:</td>
          <td height="50"><%for k = 0 to arHuyenValue(i)
			Response.Write("<img src=""../../images/TKeNgang.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
            &nbsp; <%=GetNumeric(arHuyenValue(i),0)%> <br>
            <%for k = 0 to round(arFHuyen(i)/100000)
			Response.Write("<img src=""../../images/TKeNgangDo.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
            &nbsp; <%=round(arFHuyen(i)/1000000,3)%> </td>
          <td>&nbsp;</td>
        </tr>
        <%
			end if
		next%>
        <tr>
          <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Huyện</td>
          <td>Tổng đơn hàng: <%=fTongDH%><br>
            Tổng tiền: <%=Dis_str_money(fTong)%> đ </td>
          <td valign="top"> <img src="../../images/TKeNgang.jpg" width="12" height="12">Số đơn hàng<br>
            <img src="../../images/TKeNgangDo.jpg" width="12" height="12"> Doanh thu </td>
        </tr>
      </table>
	  <%else%>
	  <table border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">
        
        <tr>
          <td style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1" align="right">Huyện\ĐH</td>
          <td style="border-bottom:#000000 solid 1">&nbsp;</td>
          <td align="right" style="border-bottom:#000000 solid 1">Đơn hàng </td>
        </tr>
        <%
	   fTongDH = 0
	   	for i = 0 to iCount
			if arHuyenValue(i) >= iDonHang  then
			fTongDH	= fTongDH+arHuyenValue(i)
		%>
        <tr>
          <td style="border-right:#000000 solid 1" align="right"><%=arHuyen(i)%>:</td>
          <td height="25"><%for k = 0 to arHuyenValue(i)
			Response.Write("<img src=""../../images/TKeNgang.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
            &nbsp; <%=GetNumeric(arHuyenValue(i),0)%> </td>
          <td>&nbsp;</td>
        </tr>
        <%
			end if
		next%>
        <tr>
          <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Huyện</td>
          <td>Tổng: <%=fTongDH%></td>
          <td>&nbsp;</td>
        </tr>
      </table>
	  <table border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
        <tr>
          <td width="25" >&nbsp;</td>
          <td width="106" align="right" style="border-bottom:#000000 solid 1px;border-right:#000000 solid 1">Huyện\Doanh thu </td>
          <td width="60" style="border-bottom:#000000 solid 1">&nbsp;</td>
          <td width="54" align="right" style="border-bottom:#000000 solid 1">Triệu</td>
        </tr>
        <%
	   fTong = 0
	   	for i = 0 to iCount
		if arHuyenValue(i) >= iDonHang  then
		fTong	= fTong+arFHuyen(i)
		%>
        <tr>
          <td >&nbsp;</td>
          <td style="border-right:#000000 solid 1" align="right"><%=arHuyen(i)%>:</td>
          <td height="25"><%for k = 0 to round(arFHuyen(i)/100000)
			Response.Write("<img src=""../../images/TKeNgangDo.jpg"" width=""1"" height=""12"" align=""middle"">")
		next%>
            &nbsp; <%=round(arFHuyen(i)/1000000,3)%> </td>
          <td>&nbsp;</td>
        </tr>
        <%
			end if
		next%>
        <tr>
          <td >&nbsp;</td>
          <td align="right" style="border-right:#000000 solid 1px;" height="100" valign="bottom">Huyện</td>
          <td>Tổng: <%=Dis_str_money(fTong)%> đ</td>
          <td>&nbsp;</td>
        </tr>
      </table>
	  	<%end if%>
	  </td>
	  </tr>
	  </table>
	  <%end if%>
	  <p>&nbsp;</p></td>
  </tr>
</table>

	<%
END IF	
IF Request.form("action")<>"Search"  THEN
Call Footer()
end if
%>
</body>
</html>
