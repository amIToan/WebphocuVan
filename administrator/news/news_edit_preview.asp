<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->

<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<SCRIPT language=JavaScript1.2 src="/Scripts/news.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Dim Upload 'Su dung AspUpload
	Set Upload = Server.CreateObject("Persits.Upload")

	Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save
%>
<TABLE>
	<TBODY>
		<tr> 
          <td height="30" valign="top">
		  	<%if Trim(Upload.Form("SubTitle"))<>"" then%>
			   <font class="main_subtitle"><%=Upload.Form("SubTitle")%></font><br>
			<%End if%>
			<font size="2" face="arial"><strong><font color="856C34"><%=Upload.Form("Title")%></font></strong></font>
			<br><font class="News_Date"><%=Hour(ConvertTime(now))%>h<%=Minute(ConvertTime(now))%>'&nbsp;<%=Day(ConvertTime(now))%>/<%=Month(ConvertTime(now))%>/<%=Year(ConvertTime(now))%></font>
		  </td>
        </tr>
        <tr> 
          <td>
		  <%set smallpicture = Upload.Files("SmallPictureFileName")
		  if Not (smallpicture Is Nothing) then%>
		  	<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000080" width="100" id="AutoNumber49" <%if Upload.form("PictureAlign")="right" then%>align="right"<%else%>align="left"<%end if%>>
                <tr>
                  <td width="1%">
				  	<%set largepicture = Upload.Files("LargePictureFileName")
				  	if Not (largepicture Is Nothing) then%>
				  	<a href="javascript: openImage('<%=replace(largepicture.OriginalPath,"\","/")%>');">
						<img src="<%=smallpicture.OriginalPath%>" border="0">
					</a>
					<%Else%>
						<img src="<%=smallpicture.OriginalPath%>" border="0">
					<%End if%>
				  </td>
                </tr>
				<tr><td align="center"><font face="Arial" size="1"><%=Upload.Form("PictureCaption")%><%if Trim(Upload.Form("PictureAuthor"))<>"" then%>&nbsp;Ảnh: <%=Upload.Form("PictureAuthor")%><%End if%></font></td></tr>
              </table>
		 <%
         'Kiểm tra xem trước đã có ảnh chưa
         Elseif Upload.Form("sSmallPicture")<>"" then
         'Nếu đã có thì hiển thị
         %>
         	<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000080" width="100" id="AutoNumber49" <%if Upload.form("PictureAlign")="right" then%>align="right"<%else%>align="left"<%end if%>>
                <tr>
                  <td width="1%">
				  	<%if Upload.Form("sLargePicture")<>"" then%>
				  	<a href="javascript: openImage('<%=NewsImagePath%><%=Upload.Form("sLargePicture")%>');">
						<img src="<%=NewsImagePath%><%=Upload.Form("sSmallPicture")%>" border="0">
					</a>
					<%Else%>
						<img src="<%=NewsImagePath%><%=Upload.Form("sSmallPicture")%>" border="0">
					<%End if%>
				  </td>
                </tr>
				<tr><td align="center"><font face="Arial" size="1"><%=Upload.Form("PictureCaption")%><%if Trim(Upload.Form("PictureAuthor"))<>"" then%>&nbsp;Ảnh: <%=Upload.Form("PictureAuthor")%><%End if%></font></td></tr>
              </table>
		 <%End if%>
            <p align="justify"><font size="2" face="arial">
				<%
					content=Replace(Upload.Form("bodyx"),"&shy;","")
					content=Replace(content,"&#39;","'")
					Response.Write content
				%>
			</font></p>
            <p align="right"><strong><font size="2" face="Arial, Helvetica, sans-serif">
				<%if Trim(Upload.Form("Author"))<>"" and Trim(Upload.Form("Source"))<>"" then
					Response.Write Upload.Form("Author") & "&nbsp;-&nbsp;" & Upload.Form("Source")
				elseif Trim(Upload.Form("Source"))<>"" then
					Response.Write Upload.Form("Source") 
				elseif Trim(Upload.Form("Author"))<>"" then
					Response.Write Upload.Form("Author")
				end if%>
			</font></strong></p>
            </td>
        </tr>
		<tr><td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng cửa sổ</a></font></td></tr>
	</TBODY>
</TABLE>
<%Set Upload=nothing%>
</body>
</html>