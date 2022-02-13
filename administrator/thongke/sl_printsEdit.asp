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
	Response.Write("<b>.&nbsp;"&strNameCategory&"</b>")
	Dim arSubCat
	arSubCat = getChilds(CatID)
	if iMumberChild = 0 then
		DisNews(CatID)
	else
		For t = 0 to iMumberChild -1			
			Response.Write("<br><br><b>-&nbsp;"&arSubCat(1,t)&"</b><br>")
			DisNews(arSubCat(0,t))		
		next
	end if
%>

<%
sub DisNews(CategoryTemp)
		sql="SELECT * "
		sql=sql & " FROM V_News n"
		sql=sql & " WHERE  n.CategoryID ='" & CategoryTemp & "'"
		'sql=sql & " and (DATEDIFF(dd, n.CreationDate '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "
		sql=sql & " ORDER BY n.NewsId desc"
		Set rsNews = Server.CreateObject("ADODB.Recordset")
		stt=1
		rsNews.open sql,con,1
%>
<form name="DanhMuc" action="UpDateGiaChuyenMuc.asp" method="post" target="_blank"> 	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  style="border:#000099 solid 1" class="CTxtContent">
	<tr>
		<td width="3%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(1,1,1,1)%>"><b>STT</b></td>
		<td width="46%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên</strong></td>
		<td width="20%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>NXB</strong></td>
		<td width="7%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong><strong>Gía bìa </strong></strong></td>
		<td width="6%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>"><strong>Giá giảm </strong></td>
		<td width="4%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>"><strong>CK %</strong></td>
		<td width="14%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>"><strong>Trạng thái </strong></td>
	</tr>
	
<%
		do while not rsNews.eof 
			Giabia	=	CLng(rsNews("giabia"))
			Gia		=	CLng(rsNews("Gia"))
			iVAT	=	round((Giabia-Gia)*100/Giabia)
%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		  <%=stt%>
		  <input name="NewsID<%=stt%>" value="<%=rsNews("NewsID")%>" type="hidden">		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  	strTemp		=	LCase(rsNews("Title"))
			str			=	left(strTemp,1)
			str			=	Ucase(str)
			strTemp		=	str	+ Right(strTemp,len(strTemp)-1)
			Response.Write(strTemp)
		  %><br>
		 <font class="CSubTitle"><i>Tác giả</i>: <%=rsNews("tacgia")%></font>		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  		  	<%=rsNews("nxb")%><br>
		  <font class="CSubTitle">
		  	 <%
		  	strTemp	=	right(rsNews("namxuatban"),1)
			if strTemp <> "" and strTemp <> "0" then
				Response.Write("Quý"&strTemp&"/")
			end if
		  	strTemp	=	Left(rsNews("namxuatban"),4)
			Response.Write(strTemp)
			%>
		  </font>		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
		  <%=Dis_str_money(Giabia)%>
		  <input name="Giabia<%=stt%>" value="<%=Giabia%>" type="hidden"> </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
		  <input name="Gia<%=stt%>" type="text" class="CTextBoxUnder" id="Gia<%=stt%>" value="<%=Dis_str_money(Gia)%>" size="5" style="text-align:right;" onKeyUp="javascript: DisMoneyThis(this); TinhVAT(<%=stt%>);">		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">			
		  <input name="iVAT<%=stt%>" type="text" class="CTextBoxUnder" value="<%=iVAT%>" size="2" style="text-align:center" id="iVAT" onBlur="checkIsNumber(this)" onKeyUp="javascript: TinhGia(<%=stt%>);">		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" >
		  <%=rsNews("Tinhtrang")%><br>
		  <font class="CSubTitle">
		  <%
		  if rsNews("Het") = 1 then
		  	Response.Write("Hết")
		  else
		  	Response.Write("Còn")
		  end if
		  %>
		  </font>		  </td>
	  </tr>

<%
			stt	=	stt	+1
			rsNews.movenext
		loop
%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">&nbsp;</td>
    </tr>
		<tr>
		  <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
		  <td bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td align="right" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td align="right" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">
		  <input name="iVATAll" type="text" class="CTextBoxUnder" value="0" size="2" style="text-align:center" onBlur="checkIsNumber(this)" onKeyUp="javascript: Allvat();">
		  </td>
		  <td align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
    </tr>
		<tr>
		  <td colspan="7" align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		  <input type="hidden" value="<%=stt-1%>" name="stt">
		  <input type="submit" name="Submit" value="cập nhật">
	      <input type="reset" name="Submit2" value="  Hủy  "></td>
    </tr>
</table>
</form>
<%
	end sub
%>
<br>
</body>
</html>
<script language="javascript">
function TinhVAT(STT)
{
	GiaBia = 0
	Gia	=	0
	strTemp = "GiaBia	=	document.DanhMuc.Giabia"+STT+".value;";
	eval(strTemp);	
	strTemp = "Gia	=	document.DanhMuc.Gia"+STT+".value;";
	eval(strTemp);

	for(k=1;k<=5;k++)
	{
		Gia = Gia.replace(",","")
	}
	iCK = GiaBia-Gia;
	iCK	= iCK/GiaBia
	iCK = Math.round(100*iCK)	
	strTemp = "document.DanhMuc.iVAT"+STT+".value	=	iCK;";
	eval(strTemp);
}
function TinhGia(STT)
{
	GiaBia = 0
	Gia	=	0
	strTemp = "GiaBia	=	document.DanhMuc.Giabia"+STT+".value;";
	eval(strTemp);	
	strTemp = "iCK	=	document.DanhMuc.iVAT"+STT+".value;";
	eval(strTemp);
	Gia	=	GiaBia	- GiaBia*iCK/100	
	strTemp = "document.DanhMuc.Gia"+STT+".value	=	Gia;";
	eval(strTemp);
	strTemp=	"DisMoneyThis(document.DanhMuc.Gia"+STT+");";
	eval(strTemp);
}

function Allvat()
{
	iSTT	=	document.DanhMuc.stt.value;
	iVATall	=	document.DanhMuc.iVATAll.value;
	for(i=1;i<=iSTT;i++)
	{
		strTemp = "document.DanhMuc.iVAT"+i+".value	=	iVATall;";
		eval(strTemp);
	}
	
}

</script>
