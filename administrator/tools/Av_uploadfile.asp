<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Call AuthenticateWithRole(AudioVideoCategoryId,Session("LstRole"),"ap")
	Av_id=GetNumeric(Request.QueryString("id"),0)
	if Av_id=0 then
		response.redirect("/administrator/")
		response.end
	end if
	Av_Type=GetNumeric(Request.QueryString("Av_Type"),0)
	if Av_Type<0 or Av_Type>4 then
		response.redirect("/administrator/")
		response.end
	end if
	
	Dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="Select Av_Title from AudioVideo where Av_id=" & Av_id
	rs.open sql,con,1
		Av_Title=Trim(rs("Av_Title"))
	rs.close
	set rs=nothing
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=AudioVideoPathUpload%>" method="post" enctype="multipart/form-data" name="fUploadFile">
  <table width="100%" border="0" cellspacing="2" cellpadding="1">
    <tr align="center" valign="middle"> 
      <td colspan="2" valign="bottom">
      	<p>&nbsp;</p>
      	<font size="3" face="Arial, Helvetica, sans-serif"><strong>Upload file</strong></font><br>
      	<font size="2" face="Arial, Helvetica, sans-serif"><strong><font color="red">*Lưu ý:</font></strong> Chỉ nhận 5 loại file <b>*.wmv</b>, <b>*.wma</b>, <b>*.avi</b>,<b>*.mp3</b> và <b>*.rm</b>, dung lượng&lt;30MB</font><br><br><br>
      </td>
    </tr>
    <tr> 
      <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><b>Tiêu đề file:</b> <%=Av_Title%></font><br><br></td>
    </tr>
    <tr> 
      <td align="center" colspan="2">
      	<font size="2" face="Arial, Helvetica, sans-serif">File:</font>
      	<input name="Av_Path" type="file" id="Av_Path" size="21">
      	<input type="submit" name="Submit" value="Upload">
      	<input type="hidden" name="av_id" value="<%=av_id%>">
      	<input type="hidden" name="Av_Type" value="<%=Av_Type%>">
      	<input type="hidden" name="LstRole" value="<%=Session("LstRole")%>">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
