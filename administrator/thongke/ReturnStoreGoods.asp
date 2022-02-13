<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%
	Ngay1			=	GetNumeric(Request.form("Ngay1"),0)
	Thang1			=	GetNumeric(Request.form("Thang1"),0)
	Nam1			=	GetNumeric(Request.form("Nam1"),0)
	Ngay2			=	GetNumeric(Request.form("Ngay2"),0)
	Thang2			=	GetNumeric(Request.form("Thang2"),0)
	Nam2			=	GetNumeric(Request.form("Nam2"),0)
	strSelSearch	=	Trim(Request.Form("selSearch"))
	iOrderBy		=  	Clng(Request.Form("RaOderBy"))

	iMaorTenSach	=	Clng(Request.Form("selMaorTenSach"))
	if iMaorTenSach = 2 then
		ProviderID	=	GetNumeric(Request.Form("selProvider"),0)
	else
		strMaorTenSach	=	Trim(Request.Form("txtMaOrTensach"))
	end if

	iCheckCategory		=	GetNumeric(Request.form("iCheckCategory"),0)
	iOnlyInventory		=	GetNumeric(Request.form("iInventory"),0)
	iBookInventory		=	GetNumeric(Request.form("iBookInventory"),0)
	iDetailInventore	=	GetNumeric(Request.form("iDetailInventore"),0)
	
	FromDate=Thang1&"/"&Ngay1&"/"&Nam1
	ToDate=Thang2&"/"&Ngay2&"/"&Nam2

		
%>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td width="48%"><div align="center"><img src="../../images/logoxseo.gif" width="100" height="41"></div></td>
    <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
    <em>ĐT: <%=soDT%>  - Email: info@xseo.com</em></td>
  </tr>
  <tr>
    <td><div align="center"><strong><%=TenGD%></strong></div></td>
    <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%>  </em></div></td>
  </tr>
</table>
<br>
<br>
<div align="center" class="CTieuDe">
  THỐNG KÊ TỒN KHO
</div>
 <div align="center" class="CTxtContent">Ngày thống kê <%=Day(now)%>/<%=Month(Now)%>/<%=year(Now)%></div>

<br>
<%if iMaorTenSach = 2  then%>
<font class="CTieuDeNho" style="background-color:#FFCCFF">Nhà cung cấp:<%=getProviderFormID(ProviderID)%></font>
<%end if%>
<%
if iMaorTenSach = 0 or iMaorTenSach = 1 or iMaorTenSach=2 then	
	select case iMaorTenSach 
		case 0
			sql = "SELECT  * FROM inputProduct where   AccountingSigna<>0 and StoreSigna<>0 and CreaterSigna<>0 "
			sql=sql + " AND (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) "
			sql=sql + " AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0)"			
		case 1
			sql = "SELECT  * FROM inputProduct where   AccountingSigna<>0 and StoreSigna<>0 and CreaterSigna<>0 "
			sql = sql + " and Maso like N'%"& strMaorTenSach &"%'"
		case 2
			sql = "SELECT  * FROM inputProduct where  AccountingSigna<>0 and StoreSigna<>0 and CreaterSigna<>0 "
			sql = sql + " and ProviderID = "&ProviderID
			sql=sql + " AND (DATEDIFF(dd, DateTime, '" & FromDate & "') <= 0) "
			sql=sql + " AND (DATEDIFF(dd, DateTime, '" & ToDate & "') >= 0)"				
	end select
	
	if iOrderBy = 1 then
		sql=sql + " ORDER BY DateTime"
	elseif iOrderBy=2 then
		sql=sql + " ORDER BY ORDER BY DateTime Desc"
	end if

	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1
	iDisInvoice = 0
	if rs.eof then
		Response.Write("<p><font class=""CTieuDeNhoNho""><img src=""../../images/icons/OK.png"" height=16 width=16 align=absmiddle> Không có dữ liệu</p> ")	
	end if

	iInTotal 		=0
	iOutTotal 		= 0
	iReturnTotal	=	0
	fInventoryTotal		= 0
	iInventoryTotal	=0
	fCoverInventoryTotal=0	
	isTitle	= true
	iSTT=0
	iInvNo	=	1	
	do while not rs.eof
		bDisInvoice = false 
		iDisInvoice = iDisInvoice + 1
		Maso				=	rs("Maso")
		inProductID			=	rs("inProductID")
		WorkerMuaHangID		=	rs("WorkerMuaHangID")
		AccountingID	=	rs("AccountingID")	
		WorkerThanhToanID	=	rs("WorkerThanhToanID")
		DateTime			=	rs("DateTime")
%>
<span id="DisInvoice<%=iDisInvoice%>"	 style="display:">
<table width="100%"  border="0" cellpadding="0" cellspacing="0"  class="CTxtContent">
<%if iBookInventory <> 1 then%>
		<tr >
		  <td colspan="10" class="CTxtContent" style="<%=setStyleBorder(1,1,1,1)%>" >

<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#FFFFCC" class="CTxtContent">
                <tr>
                  <td ><img src="../../images/icons/door_in.png" align="absmiddle"> <%=iInvNo%>. <em>Hóa đơn</em>: <b><a href="Report_SoHD.asp?inProductID=<%=inProductID%>" title="HÓA ĐƠN NHẬP HÀNG"><%=Maso%></a></b> &nbsp;&nbsp;&nbsp;<em>Ngày nhập</em>:<%=GetFullDate(DateTime)%> &nbsp;&nbsp;&nbsp; <%if iMaorTenSach <> 4 then%>
                    <em>Nhà cung cấp</em>: <b> <%=getProviderFormID(rs("ProviderID"))%></b>
                  <%end if%></td>
                </tr>
            </table>			 </td>
  </tr>
<%end if%>  
<%
	strProd="SELECT Product.ProductID, Product.inProductID,Product.NewsID as ID, Product.Unit, Product.Number, Product.Giabia, Product.Price, Product.VAT, SanPhamNhap.Title, SanPhamNhap.Tacgia FROM Product INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID where  inProductID="&inProductID
		set rsProd=Server.CreateObject("ADODB.Recordset")
		rsProd.open strProd,Con,1
		
if isTitle = true then		
%>  
		<tr align="center" valign="middle"    >
		  <td colspan="10"  class="CSubTitle" style="<%=setStyleBorder(0,0,0,1)%>" align="right">N: số lượng nhập; X: số lượng xuất;</td>
  </tr>
		<tr align="center" valign="middle" bgcolor="#FFFFFF"  >
			<td  class="CTieuDeNhoNho" style="<%=setStyleBorder(1,1,0,1)%>">TT</td>
			<td class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Tên sản phẩm</td>			
		    <td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Chuyên mục </td>
			<%if iBookInventory = 1 and iOnlyInventory=1 then%>
			<td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Ngày nhập </td>
			<%end if%>
			<td class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Giá bìa </td>
		    <td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Giá nhập </td>
		    <td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">N </td>
			<td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>"> X </td>
			<td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>"> Trả </td>
			<td  class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,0,1)%>">Tồn			</td>
      </tr>
<%		
	end if	 
		Do while not rsProd.eof
			iNumOutStore  = GetNumInvoiceOutStore(rsProd("ProductID"))
			iNumInStore	  =	rsProd("Number")
			iNumReturnStore=GetNumInvoiceReturnProvice(rsProd("ProductID"))
			iNumInventory= iNumInStore - iNumOutStore - iNumReturnStore
			iOutTotal	=	iOutTotal+iNumOutStore
			iInTotal	=	iInTotal+iNumInStore
			iReturnTotal	=	iReturnTotal+iNumReturnStore
			iInventoryTotal	=	iInventoryTotal +iNumInventory
			if  iNumInventory > 0 then
				isTitle	= false
				iSTT=iSTT+1
				fCoverInventoryTotal = fCoverInventoryTotal + rsProd("Giabia")*iNumInventory
				fInventoryTotal = fInventoryTotal + rsProd("Price")*iNumInventory	
				
				bDisInvoice	=	true

%>		<tr >
			<td align="center" width="5%" style="<%=setStyleBorder(1,1,0,1)%>"><%=iSTT%></td>
			<td   style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProd("Title")%></td>	
		 	<td   width="12%" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetRootCategory(rsProd("ID"))%></td>
			<%if iBookInventory = 1 and iOnlyInventory=1 then%>
			<td   width="8%" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetFullDate(DateTime)%></td>
			<%end if%>
			<td   width="7%" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsProd("Giabia"))%></td>
		 	<td  width="7%" style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(rsProd("Price"))%></td>
    		<td width="3%" align="center"  style="<%=setStyleBorder(0,1,0,1)%>"><%=iNumInStore%></td>
			<td width="3%" align="center"  style="<%=setStyleBorder(0,1,0,1)%>">
		  <%=iNumOutStore%></td>
		    <td width="3%" align="center"  style="<%=setStyleBorder(0,1,0,1)%>"><%=iNumReturnStore%></td>
	      <td width="3%" align="Center" style="<%=setStyleBorder(0,1,0,1)%>">
	  	  <%=iNumInventory%>		  </td>
	    </tr>

		<%
		end if		
		rsProd.movenext
		Loop
		set rsProd = nothing
		%>
</table >
<%if iBookInventory <> 1 then Response.Write("<br>") end if%>
</span>

<%
	
	if bDisInvoice = false then
		if iOnlyInventory <> 1 then
		Response.Write("<p><font class=""CTxtContent""><img src=""../../images/icons/OK.png"" height=16 width=16 align=absmiddle> Hóa đơn <b><a href=""Report_SoHD.asp?inProductID="&inProductID&""" title=""HÓA ĐƠN NHẬP HÀNG"">"&Maso&"</a></b> &nbsp;&nbsp;&nbsp; Ngày: "&GetFullDate(DateTime))
		if iMaorTenSach <> 4 then
          Response.Write(" &nbsp;&nbsp;&nbsp; Nhà cung cấp: <b>"&getProviderFormID(rs("ProviderID"))&"</b>")
		end if
		Response.Write(" đã xuất hết</p>")
		end if
		Response.Write("<script language=""javascript"">document.getElementById(""DisInvoice"&iDisInvoice&""").style.display=""none"";</script>")
	else
		iInvNo		=	iInvNo +1
	end if

	rs.movenext
	Loop   
	set rs = nothing

elseif iMaorTenSach = 3 or iMaorTenSach = 4 or iMaorTenSach=5 then

	select case iMaorTenSach 
		case 3
			sql = "SELECT * FROM SanPhamNhap"
		case 4
			sql="SELECT  * from SanPhamNhap "
			sql = sql + " where idSanPham = N'"& strMaorTenSach &"'"
		case 5
			sql="SELECT  * from SanPhamNhap "
			sql = sql + " where Title like N'%"& strMaorTenSach &"%'"
	end select
	
	if iOrderBy = 1 then
		sql=sql + " ORDER BY Title"
	elseif iOrderBy=2 then
		sql=sql + " ORDER BY Title DESC"
	end if
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1
	if rs.eof then
		Response.Write("<p><font class=""CTieuDeNhoNho""><img src=""../../images/icons/OK.png"" height=16 width=16 align=absmiddle> Không có dữ liệu</p> ")	
	end if
	iInTotal 		=0
	iOutTotal 		= 0
	iReturnTotal	=	0
	fInventoryTotal		= 0
	iInventoryTotal	=0
	fCoverInventoryTotal=0	
	iSTT=1
%>
	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  class="CTxtContent">
		<tr align="center" valign="middle" bgcolor="#FFFFFF">
			<td width="3%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(1,1,1,1)%>">TT</td>
			<td width="8%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Mã </td>
           <td  bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tên sản phẩm</td>
		   	<%if iDetailInventore <> 0 then%>
  			<td  bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Ngày nhập </td>
			<%end if%>
			<td width="5%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Giá bìa </td>
			<td width="3%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">N </td>
			<td width="3%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">X </td>
			<td width="4%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Trả </td>
			<td width="4%" bgcolor="#FFFFCC" class="CTieuDeNhoNho" style="<%=setStyleBorder(0,1,1,1)%>">Tồn			</td>
      </tr>
	  
		<%
	Do while not rs.eof
			iNumOutStore  = GetNumOutGoodsStore(rs("NewsID"))
			iOutTotal	=	iOutTotal+iNumOutStore
			iNumInStore	  =	GetNumInGoodsStore(rs("NewsID"))
			iInTotal	=	iInTotal+ iNumInStore
			iNumReturnStore=GetNumReturnGoods(rs("NewsID"))
			iReturnTotal=iReturnTotal+iNumReturnStore
			iNumInventory= iNumInStore - iNumOutStore - iNumReturnStore
	if iNumInventory> 0 then
		iInventoryTotal=iInventoryTotal+iNumInventory
		fCoverInventoryTotal = fCoverInventoryTotal + GetDetailInput(rs("NewsID"),2,0)
		fInventoryTotal			=	fInventoryTotal+ GetDetailInput(rs("NewsID"),1,0)
%>		<tr >
		  <td align="center" width="3%" style="<%=setStyleBorder(1,1,0,1)%>"><%=iSTT%></td>	
	
			<td  align="Left" width="6%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("idsanpham")%></td>
			<td  align="Left" width="47%" style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Title")%></td>
			<%if iDetailInventore <> 0 then%>
	 	 	<td style="<%=setStyleBorder(0,1,0,1)%>">
			<%
			if iDetailInventore = 1 then	
				fInventoryTotal=fInventoryTotal+ GetDetailInput(rs("NewsID"),1,1)
			else
				fInventoryTotal=fInventoryTotal+ GetDetailInput(rs("NewsID"),4,2)
			end if
			%></td>
			<%end if%>
			<td width="5%"  align="Right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(GetCoverPriceNews(rs("NewsID")))%></td>
			<td  align="center" width="3%" style="<%=setStyleBorder(0,1,0,1)%>"><%=iNumInStore%></td>
			<td  align="Right" width="3%" style="<%=setStyleBorder(0,1,0,1)%>">
		  <%=iNumOutStore%></td>
		    <td  align="Right" width="4%" style="<%=setStyleBorder(0,1,0,1)%>"><%=iNumReturnStore%></td>
	      <td  align="Center" width="4%" style="<%=setStyleBorder(0,1,0,1)%>">
	  	  <%=iNumInventory%>		  </td>
	    </tr>
		<%

		iSTT=iSTT+1
		end if
		rs.movenext
		Loop
		set rs=nothing
		%>
</table>
<%
end if
 if iInventoryTotal > 0 then%>
<br>

<table   border="0" cellpadding="2" cellspacing="2" bgcolor="#FFCCFF" class="CTxtContent" style="border:#CCCCCC solid 1px;">
  <tr>
    <td colspan="2"  align="center" class="CTieuDeNho" style="<%=setStyleBorder(0,0,0,1)%>">TỔNG</td>
  </tr>
  <tr>
    <td style="<%=setStyleBorder(0,0,0,1)%>">Tổng  nhập: </td>
    <td style="<%=setStyleBorder(0,0,0,1)%>" align="right"><%=iInTotal%></td>
  </tr>
<%if iOutTotal> 0 then%>  
  <tr>
    <td style="<%=setStyleBorder(0,0,0,1)%>">Tổng  xuất: </td>
    <td style="<%=setStyleBorder(0,0,0,1)%>" align="right"><%=iOutTotal%></td>
  </tr>
<%end if%>  
<%if iReturnTotal> 0 then%>  
  <tr>
    <td style="<%=setStyleBorder(0,0,0,1)%>" >Tổng  trả kho: </td>
    <td style="<%=setStyleBorder(0,0,0,1)%>" align="right"><%=iReturnTotal%></td>
  </tr>
<%end if%>
  <tr>
    <td style="<%=setStyleBorder(0,0,0,1)%>">Tổng tồn: </td>
    <td  align="right" style="<%=setStyleBorder(0,0,0,1)%>"><%=iInventoryTotal%></td>
  </tr>
  <tr>
    <td style="<%=setStyleBorder(0,0,0,1)%>">Tổng nguyên giá tồn: </td>
    <td style="<%=setStyleBorder(0,0,0,1)%>" align="right"><%=Dis_str_money(fCoverInventoryTotal)%></td>
  </tr>
  <tr>
    <td >Tổng giá nhập tồn: </td>
    <td  align="right"><%=Dis_str_money(fInventoryTotal)%></td>
  </tr>

</table>
<%
elseif iMaorTenSach = 6 then
	call ReportCategory()
else
	Response.Write("<p><font class=""CTieudenho""> Xin đề nghị chọn thống kê lại</font></p>")	
end if
%>
  <p>&nbsp;</p>
  
<%
sub ReportCategory()

	Dim rsCat
	Set rsCat=Server.CreateObject("ADODB.Recordset")
	
	'Lấy danh sách Chuyên mục được quyền hiển thị
	if Trim(session("LstCat"))="0" then
		sqlCat="SELECT	CategoryID, CategoryName, CategoryLevel, YoungestChildren " &_
			"FROM	NewsCategory where CategoryLoai=3 " &_
			"ORDER BY LanguageId DESC, CategoryOrder"
	else
		sqlCat="SELECT	CategoryID, CategoryName, CategoryLevel, YoungestChildren " &_
			"FROM	NewsCategory "
		strCat=GetListChildrenOfListCat(session("LstCat")) & " " & GetListParentOfListCat(session("LstCat"))
		if strCat<>"" then
			ArrCat=Split(" " & strCat & " ")
			j=0
			for i=1 to UBound(ArrCat)
				if IsNumeric(ArrCat(i)) then
					j=j+1
					if j=1 then
						sqlCat=sqlCat & "Where CategoryId=" & ArrCat(i)
					else
						sqlCat=sqlCat & " or CategoryId=" & ArrCat(i)
					end if
				end if
			next
		end if 'if strCat<>"" then
		sqlCat=sqlCat & " and CategoryLoai=3 ORDER BY LanguageId DESC, CategoryOrder"
	end if 'if Trim(session("LstCat"))="0" then
	
	'response.write sqlCat
	rsCat.open sqlCat,con,3

%>

<table width="900" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#000000">
  <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
    <td width="10%" class="CTieuDeNhoNho"><font size="2" face="Arial, Helvetica, sans-serif"><strong>TT</strong></font></td>
    <td width="49%" class="CTieuDeNhoNho"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tên chuyên mục</strong></font></td>
    <td width="21%" align="center" class="CTieuDeNhoNho"><font size="2" face="Arial, Helvetica, sans-serif">Số lượng </font></td>
    <td width="20%" align="center" class="CTieuDeNhoNho"><font size="2" face="Arial, Helvetica, sans-serif">Lượt 
      xem
    </font></td>
  </tr>
  <%
  'Lượt xem tin
  NewsCount_Total=0 'Tổng số
  STT=0
  Dim rsCount
  Set rsCount=Server.CreateObject("ADODB.Recordset")
  HTML=""
  sHTML=""
  Do while not rsCat.eof
  	STT=STT+1
  	if bg_color="#E6E8E9" then
  		bg_color="#FFFFFF"
  	else
  		bg_color="#E6E8E9"
  	end if
  	
  	if rsCat("CategoryId")=61 then
  	'Đếm thống kê cho Audio_Video
  		sqlCount_AV="SELECT	 COUNT(Av_id) AS Av_Total, SUM(Av_Count) AS Av_Count_Total " &_
			"FROM	AudioVideo "
	
		rsCount.open sqlCount_AV,con,3
    		Av_Total=Clng(rsCount("Av_Total"))
    		if IsNumeric(rsCount("Av_Count_Total")) then
  				Av_Count_Total=Clng(rsCount("Av_Count_Total"))
  			else
  				Av_Count_Total=0
  			end if
  		rsCount.close
		Reply_Total=0
		
  	elseif Clng(rsCat("CategoryLevel"))=1 then
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (CategoryId=" & rsCat("CategoryId") & " Or ParentCategoryID=" & rsCat("CategoryId") & ") "&_
		"		GROUP BY NewsID " &_
		") as View1"
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (v.CategoryId=" & rsCat("CategoryId") & " Or v.ParentCategoryID=" & rsCat("CategoryId") & ") " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"
		rsCount.open sqlCount_News,con,3
    		NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    		if IsNumeric(rsCount("NewsCount_Total")) then
  				NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  			else
  				NewsCount_Total=0
  			end if
  		rsCount.close
		rsCount.open sqlCount_Reply,con,3
			Reply_Total=Clng(rsCount("Reply_Total"))
		rsCount.close
	else
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (CategoryId=" & rsCat("CategoryId") & ") "&_
		"		GROUP BY NewsID " &_
		") as View1"
	
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (v.CategoryId=" & rsCat("CategoryId") & ") " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"
		rsCount.open sqlCount_News,con,3
    		NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    		if IsNumeric(rsCount("NewsCount_Total")) then
  				NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  			else
  				NewsCount_Total=0
  			end if
  		rsCount.close
		rsCount.open sqlCount_Reply,con,3
			Reply_Total=Clng(rsCount("Reply_Total"))
		rsCount.close
	end if 'if rsCat("CategoryId")=61 then
	
	'Hiển thị HTML
	
	if Clng(rsCat("CategoryLevel"))>1 then
		HTML="<tr bgcolor=""" & bg_color & """>" & vbNewline &_
		"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & STT & ".&nbsp;</font></td>" & vbNewline &_
		"	<td valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">&nbsp;&nbsp;&nbsp;-&nbsp;<a href=""ReturnCategoryInventory.asp?CatID="&rsCat("CategoryId")&""" target=""_blank"">" & rsCat("CategoryName") & "</a></font></td>" & vbNewline &_
		"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & CStr(FormatNumber(NewsOnline_Total,0)) & "&nbsp;&nbsp;</font></td>" & vbNewline &_
    	"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & CStr(FormatNumber(NewsCount_Total,0)) & "&nbsp;&nbsp;</font></td>" & vbNewline 	  
  		Response.Write("</tr>")

  	else
  		HTML="<tr bgcolor=""" & bg_color & """>" & vbNewline &_
		"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & STT & ".&nbsp;</font></td>" & vbNewline &_
		"	<td valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif""><b>&nbsp;&#8226;&nbsp;<a href=""ReturnCategoryInventory.asp?CatID="&rsCat("CategoryId")&""" target=""_blank"">" & rsCat("CategoryName") & "</a></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""ReturnSubCateInventory.asp?CatID="&rsCat("CategoryId")&""" target=""_blank"">Chi tiết chuyên mục con</a></font></td>" & vbNewline &_
		"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif""><b>" & CStr(FormatNumber(NewsOnline_Total,0)) & "</b>&nbsp;&nbsp;</font></td>" & vbNewline &_
    	"	<td align=""right"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif""><b>" & CStr(FormatNumber(NewsCount_Total,0)) & "</b>&nbsp;&nbsp;</font></td>" & vbNewline 
  		Response.Write("</tr>")

  	end if

  	response.write HTML
  	
 rsCat.movenext
 Loop

'Đếm tổng số
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY NewsID " &_
		") as View1"
	
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (DATEDIFF(dd, v.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, v.CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"

    rsCount.open sqlCount_News,con,3
    	NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    	if IsNumeric(rsCount("NewsCount_Total")) then
  			NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  		else
  			NewsCount_Total=0
  		end if
  	rsCount.close
  		
 Set rsCount=nothing
 rsCat.close
 set rsCat=nothing
 %>
 <tr>
    <td colspan="2" align="right" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif"><b>TỔNG CỘNG &nbsp;</b></font></td>
    <td align="right" bgcolor="#FFFFFF"><font size="3" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(NewsOnline_Total + Av_Total,0)%>&nbsp;&nbsp;</b></font></td>
    <td align="right" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(NewsCount_Total + Av_Count_Total ,0)%>&nbsp;&nbsp;</b></font></td>
  </tr>
</table>
<%
end sub
%>  
</body>
</html>
