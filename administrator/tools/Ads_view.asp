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
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		Ads_id=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if

	Call AuthenticateWithRole(AdvertisementCategoryId,Session("LstRole"),"ap")
	sql="SELECT	a.Ads_id, a.Ads_Title, a.Ads_Link, a.Ads_ImagesPath, a.Ads_width, a.Ads_Type, " &_
		"		a.Ads_height, a.Ads_Position, a.StatusId, a.Ads_Creator, a.Ads_CreationDate, " &_
        "		a.Ads_LastEditor, a.Ads_LastEdited, a.Ads_Note " &_
		"FROM	Ads a " &_
		"WHERE     (a.Ads_id = " & Ads_id & ")"
	Dim rs
	set rs=server.createObject("ADODB.Recordset")
	rs.open sql,con,3
%>
<html>
<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="98%" border="0" cellspacing="1" cellpadding="1" align="center">
    <tr> 
      <td><p align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="<%=rs("Ads_Link")%>" target="_blank" style="text-decoration: none"><%=rs("Ads_Title")%></a><br>
          </font></strong><font size="1" face="Arial, Helvetica, sans-serif"><em>
          (Tạo bởi: <%=rs("Ads_Creator")%> <%=GetFullDate(convertTime(rs("Ads_CreationDate")),"VN")%>
          <%if not IsNull(rs("Ads_LastEditor")) then%>
          	, Sửa: <%=rs("Ads_LastEditor")%> <%=GetFullDate(convertTime(rs("Ads_LastEdited")),"VN")%>
          <%end if%>)</em></font></p></td>
    </tr>
    <%if Clng(rs("Ads_Type"))=0 then%>
    <tr align="left"> 
      <td align="center"><a href="<%=rs("Ads_Link")%>" target="_blank"><img src="<%=NewsImagePath%><%=rs("Ads_ImagesPath")%>" border="0"></a></td>
    </tr>
    <%else%>
    <tr align="left"> 
      <td align="center">
 	   <object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" id="ShockwaveFlash3" width="<%=rs("Ads_width")%>" height="<%=rs("Ads_height")%>">
		<param name="Movie" value>
		<param name="Src" value="<%=NewsImagePath%><%=rs("Ads_ImagesPath")%>">
		<param name="Play" value="-1">
		<param name="Loop" value="-1">
		<param name="Quality" value="Best">
		<param name="SAlign" value>
		<param name="Menu" value="-1">
		<param name="Base" value>
		<param name="Scale" value="ExactFit">
		<param name="DeviceFont" value="0">
		<param name="EmbedMovie" value="-1">
		<param name="SWRemote" value>
		<embed SRC="<%=NewsImagePath%><%=rs("Ads_ImagesPath")%>"
			   width="<%=rs("Ads_width")%>" 
			   height="<%=rs("Ads_height")%>" 
			   PLAY="true" 
			   LOOP="true" 
			   QUALITY="high" 
			   PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash">
		</embed>
	</object>
      </td>
    </tr>
    <%end if%>
    <tr> 
      <td align="right">
       <table width="100%%" border="0" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
       	<tr bgcolor="#FFFFFF"> 
            <td align="left"><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Đường Link:</font></strong></td>
            <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rs("Ads_Link")%></font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td align="left"><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Vị 
              trí:</font></strong></td>
            <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=Get_Ads_Position_Name(Clng(rs("Ads_Position")))%></font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
          <td><strong><font size="2" face="Arial, Helvetica, sans-serif"> &nbsp;Trạng 
            thái:</font></strong></td>
            
          <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=Get_Ads_StatusId_Name(rs("StatusId"))%></font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            
          <td><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Chuyên 
            mục<br>
            &nbsp;hiển thị:</font></strong></td>
            <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">
            <%Dim rsCat
            Set rsCat=server.CreateObject("ADODB.Recordset")
     		sql="SELECT	d.CategoryId, d.Ads_OnlineChildren " &_
				"FROM	Ads a INNER JOIN " &_
        		"              AdsDistribution d ON a.Ads_id = d.Ads_id " &_
				"WHERE     (a.Ads_id = " & Ads_id & ") " &_
				"ORDER BY d.Ads_Order DESC"
			rsCat.open sql,con,3
			Do While not rsCat.eof
				if rsCat("Ads_OnlineChildren")=0 then
    				Response.write "&nbsp;<img src=""../images/icon_folder_locked.gif"" width=""15"" height=""15"" border=""0"" align=""absmiddle"" alt=""Chỉ hiển thị tại chuyên mục này"">"
		    	else
    				Response.write "&nbsp;<img src=""../images/icon_folder_unlocked.gif"" width=""15"" height=""15"" border=""0"" align=""absmiddle"" alt=""Hiển thị cả ở các chuyên mục con"">"
    			end if
				Select case Clng(rsCat("CategoryId"))
    				case 0
		    			response.write "&nbsp;Tất cả các chuyên mục<br>"
    				case else
	    				Response.write GetListParentCatNameOfCatId(rsCat("CategoryId")) & "<br>"
    			End select
			rsCat.movenext
			Loop
			rsCat.close
			set rsCat=nothing%>
            </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            
          <td><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Ghi 
            chú:</font></strong></td>
            
          <td align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
            &nbsp;<%=rs("Ads_Note")%></font></td>
          </tr>
        </table></td>
    </tr>
	 <tr> 
      <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng 
        cửa sổ</a></font></td>
    </tr>
  </table>
<%
	rs.close
	set rs=nothing
	con.close
	set con=nothing
%>
</body>
</html>
