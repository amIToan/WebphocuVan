<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->

<%
	CatID	=	Clng(Trim(Request.QueryString("CatID")))
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
        <td width="48%"><div align="center"><img src="../../images/logoxseo.png" width="100" height="41"></div></td>
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
	sql = "SELECT  CategoryName, YoungestChildren"
	sql = sql + " FROM NewsCategory"
	sql = sql + " WHERE CategoryID = "&CatID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1
	iMumberChild=0
	if not rs.eof then
		iMumberChild	=	Clng(rs("YoungestChildren"))
		strNameCategory	=	rs("CategoryName")	
	end if
	Response.Write("<font class=""CTieuDeNho"">+&nbsp;"&strNameCategory&"</font>")
	'Dim arSubCat
	'arSubCat = getChilds(CatID)
	if iMumberChild > 0 then
		sql="SELECT * "
		sql=sql & " FROM V_News n"
		sql=sql & " WHERE  n.ParentCategoryID ='" & CatID & "'"
		sql=sql & " ORDER BY n.Title"		
	else
		sql="SELECT * "
		sql=sql & " FROM V_News n"
		sql=sql & " WHERE  n.CategoryID ='" & CatID & "'"
		sql=sql & " ORDER BY n.Title"
	end if
	DisNews(sql)	
%>

<%
sub DisNews(sql)

		Set rsNews = Server.CreateObject("ADODB.Recordset")
		stt=1
		rsNews.open sql,con,1
%>	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  style="border:#000099 solid 1" class="CTxtContent">
	<tr>
		<td width="3%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho"  style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
		<td width="10%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho"  style="<%=setStyleBorder(1,1,1,1)%>">Mã</td>
		<td   align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tiêu đề</td>
		<td width="5%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Gía bìa </td>
	    <td width="3%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">N</td>
	    <td width="3%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">X</td>
	    <td width="3%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">T</td>
	    <td width="3%" align="center" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tồn</td>
	</tr>
	
<%
		iInputTotal =0
		iOutputTotal = 0
		fMoneyInventory	= 0
		iInventoryTotal		=0
		iReturnTotal	=	0	
		do while not rsNews.eof 
				
			iNumOutStore  = GetNumOutGoodsStore(rsNews("NewsID"))
			iOutputTotal	=	iOutputTotal+iNumOutStore
			
			iNumInStore	  =	GetNumInGoodsStore(rsNews("NewsID"))
			iInputTotal	=iInputTotal+iNumInStore
			
			iNumReturnStore=GetNumReturnGoods(rsNews("NewsID"))
			iReturnTotal	=	iReturnTotal+iNumReturnStore
			
			iNumInventory= iNumInStore - iNumOutStore - iNumReturnStore	
			
		if iNumInventory > 0 then
			Giabia	=	Clng(rsNews("Giabia"))
			iInventoryTotal	=	iInventoryTotal + 	iNumInventory
			fMoneyInventory	=	fMoneyInventory+Giabia
%>			
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		  <%=stt%>		 		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsNews("idsanpham")%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  	strTemp		=	LCase(rsNews("Title"))
			str			=	left(strTemp,1)
			str			=	Ucase(str)
			strTemp		=	str	+ Right(strTemp,len(strTemp)-1)
			Response.Write(strTemp)
		  %> </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(Giabia)%></td>
	      <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=iNumInStore%></td>
	      <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=iNumOutStore%></td>
	      <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=iNumReturnStore%></td>
	      <td style="<%=setStyleBorder(0,1,0,1)%>" align="center"><%=iNumInventory%></td>
	  </tr>
	<%
		stt	=	stt	+1
		end if
			rsNews.movenext
		loop
%>
		<tr>
		  <td colspan="3" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,0,1)%>"><strong>Tổng</strong></td>
		  <td align="right" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(fMoneyInventory)%></td>
          <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>"><%=iInputTotal%></td>
          <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>"><%=iOutputTotal%></td>
          <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>"><%=iReturnTotal%></td>
          <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>"><%=iInventoryTotal%></td>
	</tr>
</table>
<%
	end sub
%>
<br>
</body>
</html>
