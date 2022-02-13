<%session.CodePage=65001%>
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/constant.asp"-->
<%
	Images=Trim(Request.QueryString("param1"))
	Width=Clng(Request.QueryString("param2"))
	Height=Clng(Request.QueryString("param3"))
%>
<html>
<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onLoad="window.resizeTo(<%response.write (width+35)%>,<%response.write (height+70)%>)">
<table width="100%">
  <tr>
    <td>
	  <object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" id="ShockwaveFlash3" width="<%=width%>" height="<%=height%>">
		<param name="Movie" value>
		<param name="Src" value="<%=NewsImagePath%><%=Images%>">
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
		<embed SRC="<%=NewsImagePath%><%=Images%>"
			   width="<%=width%>" 
			   height="<%=height%>" 
			   PLAY="true" 
			   LOOP="true" 
			   QUALITY="high" 
			   PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash">
		</embed>
	</object>
  </td>
</tr>
<tr><td><center><a href="javascript: window.close();"><font size="2" face="Arial, Helvetica, sans-serif">Đóng cửa sổ</font></a></center></td></tr>
</table>
</body></html>