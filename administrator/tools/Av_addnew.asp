<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Av.asp" -->
<!--#include virtual="/administrator/inc/func_Datetime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Call AuthenticateWithRole(AudioVideoCategoryId,Session("LstRole"),"ap")
	if Request.Querystring("action")="Insert" then
		sError=False
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		
		'Av_id
		Av_id=GetMaxId("AudioVideo", "Av_id", "")
		'Av_Title
		Av_Title=ReplaceHTMLToText(Upload.form("Av_Title"))
		if Av_Title="" then
			sError=True
		end if
		'Av_ImagesPath
		set Av_Images = Upload.Files("Av_ImagesPath")
		If Av_Images Is Nothing Then
			Av_ImagesPath=""
		else
		   Filetype = Right(Av_Images.Filename,len(Av_Images.Filename)-Instr(Av_Images.Filename,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" then
				sError=True
				sAv_ImagesPath="Không là file ảnh"
		   else
		   		Av_ImagesPath="Av_ImagesPath_" & Av_id & "." & Filetype
		   end if
		End If
		Av_Time=GetNumeric(Upload.form("Av_Time"),0)
		if Av_Time=0 then
			sError=True
		end if
		Av_Capacity=GetNumeric(Upload.form("Av_Capacity"),0)
		if Av_Capacity=0 then
			sError=True
		end if
		Av_Type=GetNumeric(Upload.form("Av_Type"),-1)
		if Av_Type=-1 then
			sError=True
		end if
		'Av_Author
		Av_Author=ReplaceHTMLToText(Upload.form("Av_Author"))
		'Av_Source
		Av_Source=ReplaceHTMLToText(Upload.form("Av_Source"))
		'StatusId
		StatusId=ReplaceHTMLToText(Upload.form("StatusId"))
		if statusId<>"eded" and statusId<>"apap" then
			sError=True
		end if
		'Av_Note
		Av_Note=ReplaceHTMLToText(Upload.form("Av_Note"))
		
		if not sError then
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			if Av_ImagesPath <>"" then 
				Av_Images.SaveAs Path & "\" & Av_ImagesPath
			end if
			
			if Av_Type=1 then
				Av_LinkPopUp="<a href=""javascript: Avpopup(''runvideo.asp''," & Av_id & ",''" & Video_PopupWidth & "'',''" & Video_PopupHeight & "'');"">Play Video</a>"
				Av_Path="av_" & av_id & ".wmv"
			elseif Av_Type=0 then
				Av_LinkPopUp="<a href=""javascript: Avpopup(''runaudio.asp''," & Av_id & ",''" & Audio_PopupWidth & "'',''" & Audio_PopupHeight & "'');"">Play Audio</a>"
				Av_Path="av_" & av_id & ".wma"
			elseif Av_Type=3 then
				Av_LinkPopUp="<a href=""javascript: Avpopup(''runaudio.asp''," & Av_id & ",''" & Audio_PopupWidth & "'',''" & Audio_PopupHeight & "'');"">Play</a>"
				Av_Path="av_" & av_id & ".avi"
			elseif Av_Type=4 then
				Av_LinkPopUp="<a href=""javascript: Avpopup(''runaudio.asp''," & Av_id & ",''" & Audio_PopupWidth & "'',''" & Audio_PopupHeight & "'');"">Play Audio</a>"
				Av_Path="av_" & av_id & ".mp3"	

			else
				Av_Path="av_" & av_id & ".rm"
				Av_LinkPopUp="<a href=""" & RealMediaPath & Av_Path & """>Play</a>"
			end if
			sql="INSERT INTO AudioVideo (Av_id, Av_Title, Av_Path, Av_ImagesPath, " &_
							"Av_Time, Av_Capacity, Av_LinkPopUp, Av_Type, Av_Author, " &_
							"Av_Source, StatusId, Av_Note, Av_Creator) values " &_
				"(" & Av_id &_
				",N'" & Av_Title & "'" &_
				",'" & Av_Path & "'" &_
				",'" & Av_ImagesPath & "'" &_
				"," & Av_Time &_
				"," & Av_Capacity &_
				",'" & Av_LinkPopUp & "'" &_
				"," & Av_Type &_
				",N'" & Av_Author & "'" &_
				",N'" & Av_Source & "'" &_
				",'" & StatusId & "'" &_
				",N'" & Av_Note & "'" &_
				",N'" & session("User") & "')"

			rs.open sql,con,1
			
			set rs=nothing
			set upload=nothing
			response.redirect ("Av_Uploadfile.asp?id=" & Av_id & "&Av_Type=" & Av_Type)
			response.end()
		end if
	else
		StatusId=-1
		Av_Type=-1
	end if 'Of if Request.Form("action")="Insert" then
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=Insert" method="post" enctype="multipart/form-data" name="fInsertEvent">
  <table width="100%" border="0" cellspacing="2" cellpadding="1">
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2">
      	<font size="3" face="Arial, Helvetica, sans-serif"><strong>Tạo mới Tin Audio-Video</strong></font><br>
      	<font size="2" face="Arial, Helvetica, sans-serif"><strong><font color="red">* Lưu ý:</font></strong> Bạn sẽ Upload Videoclips sau</font>
      </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Tiêu đề:</font></td>
      <td><input name="Av_Title" type="text" id="Av_Title" size="35" maxlength="100" value="<%=Av_Title%>">
        <font size="2" face="Arial, Helvetica, sans-serif">(<strong><font color="#FF0000">*</font></strong>)</font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ảnh MH:</font></td>
      <td><input name="Av_ImagesPath" type="file" id="Av_ImagesPath" size="21"><font color="#FF0000" size="1" face="Arial, Helvetica, sans-serif"><%=sAsd_ImagesPath%></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Thời lượng:</font></td>
      <td><input name="Av_Time" type="text" id="Av_Time" size="2" maxlength="5" value="<%=Av_Time%>">
      <font size="2" face="Arial, Helvetica, sans-serif"><b>'s</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      	Dung lượng: <input name="Av_Capacity" type="text" id="Av_Capacity" size="2" maxlength="5" value="<%=Av_Capacity%>"><b>&nbsp;KB</b></font>
      <font size="2" face="Arial, Helvetica, sans-serif">(<strong><font color="#FF0000">*</font></strong>)</font>
      </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Loại tin:</font></td>
      <td><Select name="Av_Type">
      		<option value="1">Window Media Video (*.wmv)</option>
      		<option value="0">Window Media Audio (*.wma)</option>
      		<option value="2">Real Media (*.rm)</option>
      		<option value="3">Avi (*.avi)</option>
      		<option value="4">Mp3 (*.mp3)</option>
      	</Select>
      	<font size="2" face="Arial, Helvetica, sans-serif">(<strong><font color="#FF0000">*</font></strong>)</font>
      </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Tác giả:</font></td>
      <td><input name="Av_Author" type="text" id="Av_Author" size="35" maxlength="100" value="<%=Av_Author%>"></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Nguồn:</font></td>
      <td height="30" valign="middle"><input name="Av_Source" type="text" id="Av_Source" size="35" maxlength="100" value="<%=Av_Source%>"></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Trạng thái</font></td>
      <td valign="middle"><%Call List_Av_StatusId(StatusId,"--L&#7921;a ch&#7885;n tr&#7841;ng th&#225;i--")%><font size="2" face="Arial, Helvetica, sans-serif">(<strong><font color="#FF0000">*</font></strong>)</font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ghi chú:</font></td>
      <td height="30" valign="middle"><input name="Av_Note" type="text" id="Av_Note" size="35" maxlength="100" value="<%=Av_Note%>"></td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
		<input type="submit" name="Submit" value="Tạo mới">
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
