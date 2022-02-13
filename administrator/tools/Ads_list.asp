<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_Ads.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Call AuthenticateWithRole(AdvertisementCategoryId,Session("LstRole"),"ap")
	if IsNull(Request.form("CatId")) or Request.form("CatId")="" then
		Cat="AllBanner"
		CatId=0
	else
		Cat=ReplaceHTMLToText(Request.form("CatId"))
		CatId=GetNumeric(Cat,0)
	end if
	'Response.Write(Cat & ":" & CatId)
	if IsNull(Request.form("StatusId")) or Request.form("StatusId")="" then
		StatusId="-1"
	else
		StatusId=ReplaceHTMLToText(Request.form("StatusId"))
	end if
	
	if IsNull(Request.form("Ads_Position")) or Request.form("Ads_Position")="" then
		Ads_Position=-1
	else
		Ads_Position=GetNumeric(Request.form("Ads_Position"),-1)
	end if

	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	if 	Request.Form("Ads_Action")="Down" and IsNumeric(Request.Form("Ads_Id")) and IsNumeric(Request.Form("Ads_Order")) and (Cat="AllCat" or Clng(CatId)>0) and Ads_Position>=0 and StatusId="apap" then
		Ads_Order=Clng(Request.Form("Ads_Order"))
		Ads_Id=Clng(Request.Form("Ads_Id"))

		'Tim Ads_Order ngay duoi Ads truyen va`o
		sql="SELECT	TOP 1 ad.Ads_Order " &_
			"FROM	Ads a INNER JOIN " &_
            "			AdsDistribution ad ON a.Ads_id = ad.Ads_id " &_
			"WHERE	(ad.Ads_Order < " & Ads_Order & ") AND " &_ 
					"(a.Ads_Position = " & Ads_Position & ") AND (a.StatusId = 'apap') "
		if Cat="AllCat" then
			sql=sql & "AND (ad.CategoryId = 0) "
		else
			sql=sql & "AND (ad.CategoryId = " & CatId & ") "
		end if

		sql=sql & "ORDER BY ad.Ads_Order DESC"
		rs.open sql,con,1
		
		if rs.eof then
			rs.close
		else
			DownOrder=Clng(rs("Ads_Order"))
			rs.close
			sql="update AdsDistribution set Ads_Order=" & Ads_Order & " where Ads_Order=" & Downorder
			rs.open sql,con,1
			sql="update AdsDistribution set Ads_Order=" & DownOrder & " where CategoryId=" & CatId & " and Ads_id=" & Ads_id
			rs.open sql,con,1
		end if
	elseif 	Request.Form("Ads_Action")="Up" and IsNumeric(Request.Form("Ads_Id")) and IsNumeric(Request.Form("Ads_Order")) and (Cat="AllCat" or Clng(CatId)>0) and Ads_Position>=0 and StatusId="apap" then
		Ads_Order=Clng(Request.Form("Ads_Order"))
		Ads_Id=Clng(Request.Form("Ads_Id"))

		'Tim Ads_Order ngay Tren Ads truyen va`o
		sql="SELECT	TOP 1 ad.Ads_Order " &_
			"FROM	Ads a INNER JOIN " &_
            "			AdsDistribution ad ON a.Ads_id = ad.Ads_id " &_
			"WHERE	(ad.Ads_Order > " & Ads_Order & ") AND " &_ 
					"(a.Ads_Position = " & Ads_Position & ") AND (a.StatusId = 'apap') "
		if Cat="AllCat" then
			sql=sql & "AND (ad.CategoryId = 0) "
		else
			sql=sql & "AND (ad.CategoryId = " & CatId & ") "
		end if

		sql=sql & "ORDER BY ad.Ads_Order DESC"
		rs.open sql,con,1
		
		if rs.eof then
			rs.close
		else
			UpOrder=Clng(rs("Ads_Order"))
			rs.close
			sql="update AdsDistribution set Ads_Order=" & Ads_Order & " where Ads_Order=" & UpOrder
			rs.open sql,con,1
			sql="update AdsDistribution set Ads_Order=" & UpOrder & " where CategoryId=" & CatId & " and Ads_id=" & Ads_id
			rs.open sql,con,1
		end if
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>    
    <SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="container-fluid">
<%
	Title_This_Page="Công cụ ->Banner quảng cáo"
	Call header()
	
%>
<div class="container-fluid">
    <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fBanner">
  <table align="center" cellpadding="0" cellspacing="0" width="770" class="w3-table w3-table-all">
    <tr> 
      <td align="right" valign="middle"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
	  	Liệt kê:&nbsp;</strong></font> </td>
      <td align="left" valign="middle" width="1">
	  	<select name="CatId" id="CatId">
          <option value="AllBanner"<%if Cat="AllBanner" then%> selected<%end if%>>Tất cả các Quảng cáo</option>
		  <option value="0" style="COLOR: Blue; background-color:#FFFFFF">Chỉ những quảng cáo nằm trong .......</option>
		  <option value="AllCat"<%if Cat="AllCat" then%> selected<%end if%>>Tất cả các chuyên mục</option>
		  <%Call List_CategoryWithoutSelectTag(CatId)%>
        </select></td>
	  <td align="left" valign="middle" width="1"><%Call List_Ads_Position(Ads_Position,"T&#7845;t c&#7843; c&#225;c v&#7883; tr&#237;")%></td>
      <td align="left" valign="middle" width="1"><%Call List_Ads_StatusId(StatusId,"T&#7845;t c&#7843; c&#225;c tr&#7841;ng th&#225;i")%></td>
      <td align="left" valign="middle" width="1"> <a href="#" onClick="javascript: checkme(fBanner.CatId.value);"><img name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0"></a> 
        <input type="hidden" name="action" value="Search"> </td>
    </tr>
  </table>
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function checkme(thisCatIdvalue)
 {
 	if (thisCatIdvalue==0)
	{
		alert("Bạn chưa chọn phạm vi liệt kê.");
		fBanner.CatId.focus();
		return false;
	}
	document.fBanner.submit();
 }
// -->
</SCRIPT>
<form action="Ads_updateOrder.asp" name="Ads_updateOrder" method="post">
<table width="998" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000" class="w3-table w3-table-all">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td width="5%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>STT</strong></font></td>
    <td width="25%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tên quảng cáo</strong></font></td>
    <td width="25%" align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Chuyên mục<br>
      Hiển thị</strong></font></td>
    <td width="5%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Trạng<br>thái</strong></font></td>
    <td width="25%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Vị trí</strong></font></td>
    <td width="5%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>View</strong></font></td>
    <td width="5%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Order</strong></font></td>
    <td width="5%"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: winpopup('Ads_addnew.asp','',450,370);"><img src="../images/icon-banner-new.gif" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;<a href="javascript: winpopup('Ads_addnew.asp','',450,370);">Mới</a></font></td>
  </tr>
<%
	Dim boo
	boo=false
	sql="SELECT a.Ads_id, a.Ads_Title, a.Ads_width, a.Ads_height, a.StatusId, a.Ads_Creator, " &_ 
		"		a.Ads_CreationDate, a.Ads_LastEditor, a.Ads_LastEdited, a.Ads_Position, " &_
        "		a.Ads_ImagesPath, a.Ads_view, a.Ads_orderSub, d.CategoryId, d.Ads_OnlineChildren, d.Ads_Order " &_
		"FROM   Ads a INNER JOIN" &_
        "			AdsDistribution d ON a.Ads_id = d.Ads_id "
		
		if StatusId<>"-1" then
			sql=sql & "WHERE a.StatusId='" & StatusId & "'"
			boo=true
		end if
		if Ads_Position<>-1 then
			if boo then
				sql=sql & " and"
			else
				sql=sql & " WHERE"
			end if
			sql=sql & " a.Ads_Position=" & Ads_Position
			boo=true
		end if
		if Cat="AllBanner" then
		elseif Cat="AllCat" then
			if boo then
				sql=sql & " and"
			else
				sql=sql & " WHERE"
			end if
			sql=sql & " (d.CategoryId = 0)"
		else
		'Liệt kê trong một chuyên mục
			LstCat=GetListParentCat(CatId)
			ArrLstCat=Split(" " & Trim(LstCat))
			if boo then
				sql=sql & " and"
			else
				sql=sql & " WHERE"
			end if
			sql=sql & " (d.CategoryId=0 or d.CategoryId=" & CatId
			For i=1 to UBound(ArrLstCat)
				if IsNumeric(ArrLstCat(i)) then
					sql=sql & " or (d.CategoryId=" & Clng(ArrLstCat(i)) & " and d.Ads_OnlineChildren=1)"
				end if
			Next
			sql=sql & ")"
		end if
	sql=sql & " ORDER BY d.Ads_Order desc"

	'strSQL = "SELECT Count(*) AS CatCount FROM " & strDbTable & "MemCat WHERE " & strDbTable & "MemCat.Cat_ID=" & intCatID & ";"
	'rsCatCount.Open strSQL, adoCon
	'intCatCount = CLng(rsCatCount("CatCount"))
	'rsCatCount.Close
	intNumOfAds=1
	NEWS_PER_PAGE=20
	PAGES_PER_BOOK=7
	rs.PageSize = NEWS_PER_PAGE
	rs.open sql,con,1
	if not rs.eof then
		Do while not rs.EOF
			intNumOfAds=intNumOfAds + 1
		rs.MoveNext
		Loop
		rs.MoveFirst

		if request.Form("page")<>"" and isnumeric(Request.Form("page")) then
			page=Clng(request.Form("page"))
		else
			page=1
		end if
		rs.AbsolutePage = CLng(page)
		stt=(page-1)* rs.pageSize + 1
		i=0
%>
<%Do while not rs.eof and i<rs.pagesize
	i=i+1
%>
  <tr> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=stt%>.&nbsp;</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&#8226;<a href="javascript: winpopup('Ads_view.asp','<%=rs("Ads_id")%>',500,350);" style="text-decoration: none"><%=rs("Ads_Title")%></a><br>
      (Tạo: <%=rs("Ads_Creator")%>, <%=GetFullDate(convertTime(rs("Ads_CreationDate")),"VN")%>
      <%if not IsNull(rs("Ads_LastEdited")) then
    		response.write ",Sửa: " & rs("Ads_LastEditor") & ", " & GetFullDate(convertTime(rs("Ads_LastEdited")),"VN")
    	end if%>)
    </font></td>
    <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">
		<%if rs("Ads_OnlineChildren")=0 then
    		Response.write "&nbsp;<img src=""../images/icon_folder_locked.gif"" width=""15"" height=""15"" border=""0"" align=""absmiddle"" alt=""Chỉ hiển thị tại chuyên mục này"">"
    	else
    		Response.write "&nbsp;<img src=""../images/icon_folder_unlocked.gif"" width=""15"" height=""15"" border=""0"" align=""absmiddle"" alt=""Hiển thị cả ở các chuyên mục con"">"
    	end if%>
    	<%Select case Clng(rs("CategoryId"))
    		case 0
    			response.write "Tất cả các chuyên mục"
    		case else
    			Response.write ""' GetListParentCatNameOfCatId(rs("CategoryId"))
    	End select%>
    </font></td>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">
		<%if rs("statusId")="apap" then%>
			<img src="../images/icon-activate.gif" width="16" height="16" border="0" align="absmiddle" alt="Đưa lên mạng">
		<%else%>
			<img src="../images/icon-deactivate.gif" width="16" height="16" border="0" align="absmiddle" alt="Không lên mạng">
		<%end if%>
    </font></td>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">
    	<%=Get_Ads_Position_Name(Clng(rs("Ads_Position")))%>
    </font></td>
	
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("Ads_view")%></font></td>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">
	
	       <select name="AdsOrder<% = rs("Ads_id") %>"><%
           
           'loop round to display the number of forums for the order select list
           For intLoop = 1 to intNumOfAds-1
		Response.Write("<option value=""" & intLoop & """ ")

			'If the loop number is the same as the order number make this one selected
			If rs("Ads_Order") = intLoop Then
				Response.Write("selected")
			End If

		Response.Write(">" & intLoop & "</option>")
           Next
           %>
       </select> 

	
	
	</font></td>
 
    <td height="20" align="center" valign="middle">
	<%if f_permission > 1 then%>
		<%if (Cat="AllCat" or Clng(CatId)>0) and Ads_Position>=0 and StatusId="apap" then%><a style="display:block" href="javascript: fOrder('Up',<%=rs("Ads_id")%>,<%=rs("Ads_Order")%>,<%=page%>)"><img style="display:block" src="../images/icon_go_up.gif" width="15" height="15" hspace="2" border="0" align="absmiddle"></a><%end if%><a style="display:block" href="javascript: winpopup('Ads_edit.asp','<%=rs("Ads_id")%>',450,370);"><img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle" hspace="2"></a><br>
		<%if (Cat="AllCat" or Clng(CatId)>0) and Ads_Position>=0 and StatusId="apap" then%><a style="display:block;" href="javascript: fOrder('Down',<%=rs("Ads_id")%>,<%=rs("Ads_Order")%>,<%=page%>)"><img style="display:block" src="../images/icon_go_down.gif" width="15" height="15" hspace="2" border="0" align="absmiddle"></a><%end if%><%if f_permission > 3 then %><a style="display:block" href="javascript: winpopup('Ads_delete.asp','<%=rs("Ads_id")%>&CatId=<%=rs("CategoryId")%>',300,150);"><img src="../images/icon_closed_topic.gif" width="15" height="15" hspace="2" border="0" align="absmiddle"></a><%end if%></font>
	<%end if%>
		</td>
  </tr>
<%stt=stt+1
rs.movenext
Loop%>
        <tr align="center" bgcolor="#FFFFFF"> 
            <td colspan="6">
			<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
			Trang: 
				<%
				if page>Clng(PAGES_PER_BOOK/2) then
					Minpage=page-Clng(PAGES_PER_BOOK/2)+1
				else
					MinPage=1
				end if
									
				Maxpage=Minpage+PAGES_PER_BOOK-1
				if Maxpage >rs.pagecount then
					Maxpage=rs.pagecount
				end if
									
				for i=Minpage to Maxpage 
					if i<>page then
					response.Write "<a href=""javascript: fSearch2(" & i & ");"">" & i & "</a>&nbsp;|&nbsp;"
					else
					response.Write "<font color=""red""><b>" & i & "</b></font>&nbsp;|&nbsp;"
					end if
				Next%>
			</font>
							   
			</td>
            <td colspan="2" >
			<input type="submit" name="Submit" value="Update Order">
			<input type="hidden" name="CatId" value="<%=CatId%>">
			</td>

        </tr>
							
<%	else 'if not rs.eof then
%>
  <tr bgcolor="#FFFFFF"> 
    <td align="left" colspan="8"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"><b>&nbsp;&nbsp;Không có dữ liệu</b></font></td>
  </tr>
<%	end if 'if not rs.eof then
	rs.close
	set rs=nothing
%>
</table>
</form>
	<form action="<%=Request.ServerVariables("Script_name")%>" method="post" name="fSearch2">
		<input type="hidden" name="CatId" value="<%=Cat%>">
		<input type="hidden" name="StatusId" value="<%=StatusId%>">
		<input type="hidden" name="Ads_Position" value="<%=Ads_Position%>">
		<input type="hidden" name="page">
		<input type="hidden" name="Ads_Action">
		<input type="hidden" name="Ads_Id">
		<input type="hidden" name="Ads_Order">
	</form>
	<script language="JavaScript">
		function fSearch2(page)
		{
			document.fSearch2.page.value=page;
			document.fSearch2.submit();
		}
		function fOrder(Ads_Action,Ads_Id,Ads_Order,page)
		{
			document.fSearch2.Ads_Action.value=Ads_Action;
			document.fSearch2.Ads_Id.value=Ads_Id;
			document.fSearch2.Ads_Order.value=Ads_Order;
			document.fSearch2.page.value=page;
			//alert ("Ads_Action:" + Ads_Action + ",Ads_Id:" + Ads_Id + ",Ads_Order: " + Ads_Order + ",page=" + page);
			document.fSearch2.submit();
		}
	</script>
</div>
    </div>
        </div>
<%Call Footer()%>
</body>
</html>
