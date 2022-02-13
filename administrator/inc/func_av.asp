<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Sub List_Av_StatusId(AvSelect,Av_StatusId_Title)%>
		<select name="StatusId" id="StatusId">
			<option value="-1"><%=Av_StatusId_Title%></option>
			<option value="apap"<%if AvSelect="apap" then%> selected<%end if%>>Đưa lên mạng</option>
			<option value="eded"<%if AvSelect="eded" then%> selected<%end if%>>Không lên mạng</option>
        </select>
<%End sub%>
<%Function Get_Av_StatusId_Name(StatusId)
	Select case StatusId
		case "eded"
			Get_AvStatusId_Name="Không lên mạng"
		case "apap"
			Get_Av_StatusId_Name="Đưa lên mạng"
		case else
			Get_Av_StatusId_Name=""
	End select
End Function%>
<%Sub AudioVideo_List(Av_Status)
	'Av_Status: - DESC: Chỉ dùng để minh họa trong bài viết (chỉ cần có quyền btv).
	'					Chỉ có một chức năng duy nhất là chèn Auudio_Video vào bài viết
	'			- EDIT: Có toàn bộ quyền đối với chuyên mục tin Audio_Video.
	Keyword=ReplaceHTMLToText(Request.form("keyword"))
	Cat=GetNumeric(Request.form("Cat"),0)
%>

<FORM action="<%=Request.ServerVariables("Script_Name")%>" method="post" name="fAV">
  <table align="center" cellpadding="0" cellspacing="0" width="98%">
    <tr> 
      <td align="right" valign="middle"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tìm kiếm:&nbsp;</strong></font></td>
      <td align="left" valign="middle" width="1"><input type="text" name="keyword" id="keyword" value="<%=Keyword%>"></td>
      <td align="left" valign="middle" width="1">
	  	<select name="Cat" id="Cat">
          <option value="0" style="COLOR: Red; background-color:#FFFFFF">--Phạm vi tìm kiếm--</option>
		  <option value="1"<%if Cat=1 then%> selected<%end if%>>&nbsp;&nbsp;Tiêu đề</option>
		  <option value="2"<%if Cat=2 then%> selected<%end if%>>&nbsp;&nbsp;Ngày (mm/dd/yyyy)</option>
		  <option value="3"<%if Cat=3 then%> selected<%end if%>>&nbsp;&nbsp;Tháng (mm/yyyy)</option>
		  <option value="4"<%if Cat=4 then%> selected<%end if%>>&nbsp;&nbsp;Năm (yyyy)</option>
		  <option value="5"<%if Cat=5 then%> selected<%end if%>>&nbsp;&nbsp;Mã Audio-Video</option>
        </select> </td>
      <td align="left" valign="middle" width="1"><a href="#" onClick="javascript: checkme(fAV.Cat.value,fAV.keyword.value);"><img name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0"></a> 
        <input type="hidden" name="action" value="Search"> </td>
    </tr>
  </table>
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function checkme(thisCatIdvalue,thisKeywordvalue)
 {
 	if (thisKeywordvalue=="")
	{
		alert("Bạn chưa nhập từ khóa tìm kiếm.");
		fAV.keyword.focus();
		return false;
	}
	
 	if (thisCatIdvalue==0)
	{
		alert("Bạn chưa chọn phạm vi tìm kiếm.");
		fAV.Cat.focus();
		return false;
	}
	document.fAV.submit();
 }
// -->
</SCRIPT>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td width="4%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>STT</strong></font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tiêu đề</strong></font></td>
   <%if Av_Status="EDIT" then%>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tác giả</strong></font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Nguồn</strong></font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Trạng thái</strong></font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Lượt xem</strong></font></td>
   <%end if%>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><%if Av_Status="EDIT" then%><a href="javascript: winpopup('Av_addnew.asp','',450,350);"><img src="../images/icon-banner-new.gif" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;<a href="javascript: winpopup('Av_addnew.asp','',450,350);">Mới</a></font><%End if%></td>
  </tr>
<%
	Dim rs
	Set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT	Av_id, Av_Title, Av_Path, Av_ImagesPath, Av_Time, Av_Capacity, Av_LinkPopUp, " &_
		"		Av_Type, Av_Author, Av_Source, StatusId, Av_Note, Av_Creator, " &_
        "		Av_CreationDate, Av_LastEditor, Av_LastEditedDate, Av_Count " &_
		"FROM	AudioVideo "
	Select case Cat
		case 1 'Theo Tiêu đề
			sql=sql & "WHERE (Av_Title LIKE N'%" & keyword & "%')"
		case 2 'Theo ngày tháng năm
			if IsDate(keyword) then
				strTmp=FormatDateTime(keyword)
				sql=sql & "WHERE DATEDIFF(dd, Av_CreationDate, '" & strTmp & "')=0"
			end if
		case 3 'Theo tháng năm
			strTmp="28/" & keyword
			if IsDate(strTmp) then
				strTmp=FormatDateTime(strTmp)
				sql=sql & "WHERE DATEDIFF(mm, Av_CreationDate, '" & strTmp & "')=0"
			end if
		case 4 'Theo năm
			strTmp="28/12/" & keyword
			if IsDate(strTmp) then
				strTmp=FormatDateTime(strTmp)
				sql=sql & "WHERE DATEDIFF(yy, Av_CreationDate, '" & strTmp & "')=0"
			end if
		case 5 'Theo mã Audio-Video
			sql=sql & "WHERE Av_id=" & keyword
	End select
	sql=sql & " ORDER BY Av_id DESC"
	'response.write sql
	'response.end
	
	NEWS_PER_PAGE=20
	PAGES_PER_BOOK=7
	rs.PageSize = NEWS_PER_PAGE
	rs.open sql,con,1
	
	if not rs.eof then
		if request.Form("page")<>"" and isnumeric(Request.Form("page")) then
			page=Clng(request.Form("page"))
		else
			page=1
		end if
		rs.AbsolutePage = CLng(page)
		stt=(page-1)* rs.pageSize + 1
		i=0
		
	  Do while not rs.eof and i<rs.pagesize
		i=i+1
%>
  <tr bgcolor="<%if i mod 2=0 then%>#FFFFFF<%else%>#E6E8E9<%end if%>"> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=stt%>.&nbsp;</font></td>
    <td align="justify"><p align="justify">
    <%if rs("Av_ImagesPath")<>"" then%>
    	<img src="<%=NewsImagePath%><%=rs("Av_ImagesPath")%>" border="0" vspace="2" align="left">
    <%end if%>
    <%if Clng(rs("Av_Type"))=0 then%>
    	<a href="<%=AudioVideoPath%><%=rs("Av_Path")%>"><img src="/administrator/images/i_audio.gif" width="14" height="14" border="0" align="absmiddle" hspace="2" alt="Play Audio"></a>
    <%elseif Clng(rs("Av_Type"))=1 then%>
    	<a href="<%=AudioVideoPath%><%=rs("Av_Path")%>"><img src="/administrator/images/i_video.gif" width="14" height="14" border="0" align="absmiddle" hspace="2" alt="Play video"></a>
    <%else%>
    	<a href="<%=RealMediaPath%><%=rs("Av_Path")%>"><img src="/administrator/images/icon_RM.jpg" width="14" height="14" border="0" align="absmiddle" hspace="2" alt="Play"></a>
    <%end if%>
    <font size="2" face="Arial, Helvetica, sans-serif"><%=rs("Av_Title")%><br>
      (Bởi: <%=rs("Av_Creator")%>, <%=GetFullDate(convertTime(rs("Av_CreationDate")),"VN")%>
      <%if not IsNull(rs("Av_LastEditor")) then
    		response.write ",Sửa: " & rs("Av_LastEditor") & ", " & GetFullDate(convertTime(rs("Av_LastEditedDate")),"VN")
    	end if%>)
    </font></td>
  <%if Av_Status="EDIT" then%>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("Av_Author")%></font></td>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("Av_Source")%></font></td>
    <td align="center">
    	<%if rs("statusId")="apap" then%>
			<img src="../images/icon-activate.gif" width="16" height="16" border="0" align="absmiddle" alt="Đưa lên mạng">
		<%else%>
			<img src="../images/icon-deactivate.gif" width="16" height="16" border="0" align="absmiddle" alt="Không lên mạng">
		<%end if%>
    </td>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("Av_Count")%></font></td>
    <td height="20" align="center" valign="middle"><a href="javascript: winpopup('Av_edit.asp','<%=rs("Av_id")%>',450,350);"><img src="/administrator/images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle" hspace="2"></a><a href="javascript: winpopup('Av_delete.asp','<%=rs("Av_id")%>',300,150);"><img src="/administrator/images/icon_closed_topic.gif" width="15" height="15" hspace="2" border="0" align="absmiddle"></a></td>
  <%else
  	if Clng(rs("Av_Type"))=1 then%>
  		<td height="20" align="center" valign="middle"><a href="javascript: onButtonClick(<%=rs("Av_id")%>,<%=rs("Av_Type")%>,<%=Video_PopupWidth%>,<%=Video_PopupHeight%>);"><img src="/administrator/images/icon_reply_topic.gif" width="15" height="15" border="0" align="absmiddle" hspace="2" alt="Insert"></a></td>
  	<%elseif Clng(rs("Av_Type"))=0 then%>
  		<td height="20" align="center" valign="middle"><a href="javascript: onButtonClick(<%=rs("Av_id")%>,<%=rs("Av_Type")%>,<%=Audio_PopupWidth%>,<%=Audio_PopupHeight%>);"><img src="/administrator/images/icon_reply_topic.gif" width="15" height="15" border="0" align="absmiddle" hspace="2" alt="Insert"></a></td>
  	<%else%>
  		<td height="20" align="center" valign="middle"><a href="javascript: onButtonClick(<%=rs("Av_id")%>,<%=rs("Av_Type")%>,0,0);"><img src="/administrator/images/icon_reply_topic.gif" width="15" height="15" border="0" align="absmiddle" hspace="2" alt="Insert"></a></td>
  	<%end if%>
  <%end if%>
    
  </tr>
<%stt=stt+1
rs.movenext
Loop%>
	<form action="<%=Request.ServerVariables("Script_name")%>" method="post" name="fSearch2">
		<input type="hidden" name="Cat" value="<%=Cat%>">
		<input type="hidden" name="keyword" value="<%=keyword%>">
		<input type="hidden" name="page">
	</form>
	<script language="JavaScript">
		function fSearch2(page)
		{
			document.fSearch2.page.value=page;
			document.fSearch2.submit();
		}
	</script>
                            <tr align="center" bgcolor="#FFFFFF"> 
                              <td colspan="7">
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
                            </tr>
<%	else 'if not rs.eof then
%>
  <tr bgcolor="#FFFFFF"> 
    <td align="left" colspan="6"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"><b>&nbsp;&nbsp;Không có dữ liệu</b></font></td>
  </tr>
<%	end if 'if not rs.eof then
	rs.close
	set rs=nothing
	
	if Av_Status="EDIT" then
	else%>
		<script language="javascript">
			function onButtonClick(Av_id,Av_Type,Av_width,Av_height)
			{
				switch (Av_Type)
				{
					case 0:
						popupTag="<a href=\"javascript:Avpopup('runaudio.asp'," + Av_id + ",'" + Av_width + "','" + Av_height + "');\">Play Audio</a>";
						break;
					case 1:
						popupTag="<a href=\"javascript:Avpopup('runvideo.asp'," + Av_id + ",'" + Av_width + "','" + Av_height + "');\">Play Video</a>";
						break;
					case 2:
						popupTag="<a href=\"rtsp://203.162.130.166:5554/mountpoint1/Av_" + Av_id + ".rm\">Play</a>";
						break;
				}
				opener.InsertNewImage(popupTag);
				window.close();
				window.opener.focus();
			}
		</script>
	<%end if
%>
</table>
<%End Sub%>