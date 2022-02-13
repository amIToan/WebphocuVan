<%Sub Picture_list()
	Keyword=ReplaceHTMLToText(Request.form("keyword"))
	Cat=GetNumeric(Request.form("Cat"),0)
%>
<FORM action="<%=Request.ServerVariables("Script_name")%>" method="post" name="fSearch">
  <table align="center" cellpadding="0" cellspacing="0" width="98%">
    <tr> 
      <td align="right" valign="middle"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tìm kiếm:&nbsp;</strong></font></td>
      <td align="left" valign="middle" width="1"><input type="text" name="keyword" value="<%=keyword%>" size="15"></td>
      <td align="left" valign="middle" width="1">
      	<select name="Cat" id="Cat">
          <option value="0" style="COLOR: Red; background-color:#FFFFFF">--------Phạm vi--------</option>
          <option value="1"<%if Cat=1 then%> selected<%end if%>>Tên của ảnh</option>
          <option value="2"<%if Cat=2 then%> selected<%end if%>>Chú thích của ảnh</option>
          <option value="3"<%if Cat=3 then%> selected<%end if%>>Tác giả ảnh</option>
          <option value="4"<%if Cat=4 then%> selected<%end if%>>Tin ảnh trang chủ </option>
          <option value="5"<%if Cat=5 then%> selected<%end if%>>Tin ảnh chuyên mục</option>
        </select>
      </td>
      <td align="left" valign="middle" width="1"><a href="#" onClick="javascript: checkme(fSearch.Cat.value,fSearch.keyword.value);"><img name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0"></a></td>
	  <td align="left" valign="middle" width="1"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>|</b></font></td>
	  <td align="left" valign="middle" width="1"><img src="../images/icon-banner-new.gif" width="16" height="16" border="0" align="absmiddle"></td>
	  <td align="left" valign="middle" width="1"><a href="javascript: winpopup('/administrator/picture/picture_addnew.asp','1',420,300);"><font size="2" face="Arial, Helvetica, sans-serif">Mới</font></a></td>
    </tr>
  </table>
  <input type="hidden" name="action" value="Search">
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function checkme(thisCatIdvalue,thisKeywordvalue)
 {
 	if (thisKeywordvalue=="")
	{
		alert("Bạn chưa nhập từ khóa tìm kiếm.");
		fSearch.keyword.focus();
		return false;
	}
	
 	if (thisCatIdvalue==0)
	{
		alert("Bạn chưa chọn phạm vi tìm kiếm.");
		fSearch.Cat.focus();
		return false;
	}
	document.fSearch.submit();
 }
			function onButtonClick(anhnho,anhto,ImagePath,border)
			{
				if (anhto!="")
				{
					imageTag ="<a href=\"javascript:openImage('" + ImagePath + anhto + "');\">";
					imageTag += "<IMG src=\"" + ImagePath + anhnho + "\" ";
				}
				else
				{
					imageTag = "<IMG src=\"" + ImagePath + anhnho + "\" ";
				}
				
				imageTag += "alt=\"Ảnh minh họa\" "; 
				imageTag += "align=\"center\" "; 
				imageTag += "border=\"" + border + "\">";
				if (anhto!="")
				{
					imageTag += "</a>";
				}
				opener.InsertNewImage(imageTag);
				window.close();
			}
// -->
</SCRIPT>
<%
	sql="SELECT	p.PictureID, p.PictureCaption, p.PictureName, p.SmallPictureFileName, " &_
		"		p.LargePictureFileName, p.PictureAuthor, p.CreationDate, p.Creator, " &_
        "       p.CategoryID, p.StatusID, p.IsHomePicture, p.IsCatHomePicture, " &_
        "		p.Approver, p.ApproverDate, c.CategoryName " &_
		"FROM	Picture p INNER JOIN " &_
        "        NewsCategory c ON p.CategoryID = c.CategoryID "
        
	if keyword="AllDatabase" then
		'List All Picture in Database
	else
		Select case Cat
			case 1
				sql=sql & "WHERE	(p.PictureName LIKE N'%" & Keyword & "%')"
			case 2
				sql=sql & "WHERE	(p.PictureCaption LIKE N'%" & Keyword & "%')"
			case 3
				sql=sql & "WHERE	(p.PictureAuthor LIKE N'%" & Keyword & "%')"
			case 4
				sql=sql & "WHERE	(p.IsHomePicture=1)"
			case 5
				sql=sql & "WHERE	(p.IsCatHomePicture=1)"
		End select
	end if
	sql=sql & " ORDER BY p.PictureID DESC"
	'response.write sql
	Dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
	
	NEWS_PER_PAGE=15
	PAGES_PER_BOOK=7
	rs.PageSize = NEWS_PER_PAGE
	rs.open sql,con,3
	if not rs.eof then
		if request.QueryString("page")<>"" and isnumeric(Request.QueryString("page")) then
			page=Clng(request.QueryString("page"))
		else
			page=1
		end if
		rs.AbsolutePage = CLng(page)
		i=0
%>
<table width="998"" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <%Do while not rs.eof and i<rs.pagesize
  	j=0%>
  <tr align="center" bgcolor="#FFFFFF">
    <%Do while not rs.eof and j<3%>
    <td width="33%" valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="2">
		<tr>
          <td colspan="4" bgcolor="#E6E8E9" height="23">
           <%if rs("IsHomePicture") then%>
            <img src="../images/icon-affiliate.gif" width="16" height="16" border="0" align="absmiddle" alt="Tin ảnh của trang chủ">
           <%end if%>
           <%if rs("IsCatHomePicture") then%>
            <img src="../images/icon-campaign.gif" width="16" height="16" border="0" align="absmiddle" alt="Tin ảnh của chuyên mục">
           <%end if%>
          <font size="2" face="Arial, Helvetica, sans-serif"><strong>
          	<%if Trim(rs("PictureName"))<>"" then
          		response.write rs("PictureName")
          	else
          		response.write "Ảnh không có tên"
          	end if%></strong></font></td>
        </tr>
        <tr> 
          <td colspan="4" height="160" align="center" valign="middle">
		  <img src="<%=NewsImagePath%><%=rs("SmallPictureFileName")%>" border="0" width="150" >
		  </td>
        </tr>
        <tr> 
          <td colspan="4"><div align="justify"><font size="2"><em><font face="Arial, Helvetica, sans-serif">
          	<%=rs("PictureCaption")%>
          	<%if Trim(rs("PictureAuthor"))<>"" then%>
          		(Ảnh - <strong><%=rs("PictureAuthor")%></strong>)
          	<%end if%>&nbsp;
          </font></em></font></div></td>
        </tr>
        <tr> 
          <td colspan="4" align="left" valign="bottom"><img src="../images/icon1.gif" width="7" height="7" border="0" align="absmiddle"><font size="2" face="Arial, Helvetica, sans-serif"><%=GetListParentCatNameOfCatId(rs("CategoryId"))%></font></td>
        </tr>
        <tr> 
          <td colspan="4" align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
            <img src="../images/icon1.gif" width="7" height="7" border="0" align="absmiddle">Tạo: 
            <%=rs("Creator")%> <font style="font-family: Times New Roman; font-size: 7pt; font-weight: normal; color:#999999;">(<%=GetFullDateTime(rs("CreationDate"))%>)</font></font></td>
        </tr>
        <tr> 
          <td colspan="4" align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
            <img src="../images/icon1.gif" width="7" height="7" border="0" align="absmiddle">Sửa: 
            <%if IsDate(rs("ApproverDate")) then%>
            <%=rs("Approver")%> <font style="font-family: Times New Roman; font-size: 7pt; font-weight: normal; color:#999999;">(<%=GetFullDateTime(rs("CreationDate"))%>)</font>
            <%end if%>
            </font></td>
        </tr>
        <tr> 
          <td width="25%" align="center" valign="bottom"><%if Trim(rs("LargePictureFileName"))<>"" then%><img src="../images/icon_profile.gif" width="15" height="15" border="0" align="absmiddle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: winpopup('/administrator/picture/picture_view.asp','<%=rs("LargePictureFileName")%>',100,200);" style="text-decoration: none">Xem</a></font><%end if%></td>
          <td width="25%" align="center" valign="bottom"><img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: winpopup('/administrator/picture/picture_edit.asp','<%=rs("pictureid")%>',420,300);" style="text-decoration: none">Sửa</a></font></td>
          <td width="25%" align="center" valign="bottom"><img src="../images/icon_closed_topic.gif" width="15" height="15" border="0" align="absmiddle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: winpopup('/administrator/picture/picture_delete.asp','<%=rs("pictureid")%>',300,150);" style="text-decoration: none">Xóa</a></font></td>
          <td width="25%" align="center" valign="bottom"><img src="../images/icon_reply_topic.gif" width="15" height="15" border="0" align="absmiddle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: onButtonClick('<%=rs("SmallPictureFileName")%>','<%=rs("LargePictureFileName")%>','<%=NewsImagePath%>','0');" style="text-decoration: none">Chèn</a></font></td>
        </tr>
      </table>
    </td>
    <%j=j+1
      i=i+1
      if j<3 then
        rs.movenext
      end if
    Loop
    for k=j to 2
    	response.write "<td width=""33%"" valign=""top"">&nbsp;</td>"
    next
    %>
  </tr>
  <%i=i+1
	if not rs.eof then
	  rs.movenext
	end if
  Loop%>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="3" align="right">
      <%Call phantrang(page,rs.pagecount,PAGES_PER_BOOK)%>
    </td>
  </tr>
</table>
<%end if
	rs.close
	set rs=nothing
End sub%>