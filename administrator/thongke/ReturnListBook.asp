<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->

<%
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	
	keyword=Trim(Request.Form("keyword"))
	if keyword ="" then
		keyword	=	Request.QueryString("keyword")
	end if
	Keyword=Replace(keyword,"'","''")
	seach_filter=Request.Form("select_filter")
	if seach_filter="" then
		seach_filter=Request.QueryString("seach_filter")
	end if
	isSapXep	=  Clng(Request.Form("isSapXep"))
	if isSapXep= 0 then
		isSapXep = 1
	end if

	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2

	StringSearch=Trim(Keyword)
	Set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * " &_
		"FROM	V_News "
	select case seach_filter
		case 1
			StrTitle	=	UCASE(StringSearch)
			sql=sql + "WHERE ({fn UCASE(Title)} like N'%"&StrTitle&"%' OR Description like N'%"&StrTitle&"%' OR Body like N'%"&StrTitle&"%')"
		case 2
			StrTitle	=	UCASE(StringSearch)
			sql= sql+ " Where ({fn UCASE(Title)} like N'%"&StrTitle&"%')"
		case 3
			sql= sql+ "Where (Body like N'%"&StringSearch&"%')"
		case 4
			sql= sql+ "Where (nxb like N'%"&StringSearch&"%')"
		case 5
			sql= sql+ "Where (tacgia like N'%"&StringSearch&"%')"
		case 6
			sql= sql+ "Where (Giabia = '"&StringSearch&"')"
		case 7
			StringTitle	=	UCASE(StringSearch)
			sql= sql+ "WHERE ({fn UCASE(Title)} like N'%"&StringTitle&"%' OR ({fn UCASE(CategoryName)} like N'%"&StringTitle&"%') or Description like N'%"&StringSearch&"%' OR Body like N'%"&StringSearch&"%' OR nxb like N'%"&StringSearch&"%' OR tacgia like N'%"&StringSearch&"%')"
		case 8
			sql= sql+ "Where (idsanpham like N'%"&StringSearch&"%')"
		case 9
			StringTitle	=	UCASE(StringSearch)				
			sql= sql+ "Where ({fn UCASE(CategoryName)} like N'%"&StringSearch&"%')"
		case 10
			StringTitle	=	UCASE(StringSearch)				
			sql= sql+ "Where ({fn UCASE(CategoryName)} like N'%"&StringSearch&"%')"
		case 11
			StringTitle	=	UCASE(StringSearch)				
			sql= sql+ "Where ({fn UCASE(Tinhtrang)} like N'%"&StringSearch&"%')"			
			
	end select
	sql= sql+ " AND (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0)"
	if isSapXep = 1 then
		sql= sql + "ORDER BY NewsID DESC"
	else
		sql= sql + "ORDER BY NewsID"
	end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
      <tr>
        <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
        <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
            <em>ĐT: <%=soDT%> - Email: info@xseo.com</em></td>
      </tr>
      <tr>
        <td><div align="center"><strong><%=TenGD%></strong></div></td>
        <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%>   </em></div></td>
      </tr>
</table>
	<br>
	<br>
	<%
		set rsNews=server.CreateObject("ADODB.Recordset")
		rsNews.open sql,con,1
		stt=0
	%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" style="border:#000099 solid 1" class="CTxtContent">
	<tr>
	  <td width="2%" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><b>STT</b></td>
	  <td width="33%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>TÊN</strong></td>
		<td width="18%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>TÁC GIẢ</strong></td>
		<td width="16%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>NXB</strong></td>
		<td width="2%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
		<td width="7%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>NĂM</strong></td>
		<td width="5%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>GIÁ BÌA</strong></td>
	    <td width="13%"  align="center" style="<%=setStyleBorder(0,1,1,1)%>">NCC</td>
	    <td width="4%"  align="center" style="<%=setStyleBorder(0,1,1,1)%>">Tồn </td>
	</tr>
	
<%
		do while not rsNews.eof 
			stt	=	stt	+1
%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  	strTemp		=	LCase(rsNews("Title"))
			str			=	left(strTemp,1)
			str			=	Ucase(str)
			strTemp		=	str	+ Right(strTemp,len(strTemp)-1)
			Response.Write(strTemp)
		  %>		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsNews("tacgia")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsNews("nxb")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  if rsNews("Het") = 1 then
		  	Response.Write("Hết")
		  else
		  	Response.Write("Còn")
		  end if
		  %></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  	strTemp	=	right(rsNews("namxuatban"),1)
			if strTemp <> "" and strTemp <> "0" then
				Response.Write("Quý"&strTemp&"/")
			end if
		  	strTemp	=	Left(rsNews("namxuatban"),4)
			Response.Write(strTemp)
			%>			</td>
		  <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsNews("giabia"))%></td>
	      <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsNews("Tinhtrang")%></td>
	      <td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetNumInventoryGoods(rsNews("NewsID"))%></td>
	  </tr>
<%
			rsNews.movenext
		loop
		set rsNews = nothing
%>
</table>
</body>
</html>
