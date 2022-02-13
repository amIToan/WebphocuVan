<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<%
	AudioVideoCategoryId=7
Set Upload = Server.CreateObject("Persits.Upload")
Upload.SetMaxSize 30000000, True 'Dat kich co upload la` 30MB
Upload.Save

	if Upload.Form("LstRole")="0ad" or Upload.Form("LstRole")="0ap" or Instr(" " & Trim(Upload.Form("LstRole")) & " "," " & AudioVideoCategoryId & "ap ")>0 then
		'Av_id
		av_id=Upload.form("av_id")
		if IsNumeric(av_id) then
			av_id=Clng(av_id)
		else
			set Upload=nothing
			response.write "Invalid Av_id"
			response.end()
		end if
		'Av_Path
		set Av_AudioVideo = Upload.Files("Av_Path")
		If Av_AudioVideo Is Nothing Then
			set Upload=nothing
			response.write "File Empty"
			response.end()
		else
		   	Filetype = Right(Av_AudioVideo.Filename,len(Av_AudioVideo.Filename)-Instr(Av_AudioVideo.Filename,"."))
		   		Av_Path="av_" & av_id & "." & Filetype
		   		Av_AudioVideo.SaveAs server.MapPath(DirectoryStored) & "\" & Av_Path
				Av_LinkPopUp="<a href=""" & DirectoryStored & Av_Path & """>Download</a>"
				
				set rs=server.CreateObject("ADODB.Recordset")
				sql="Update AudioVideo Set" &_
				" Av_Path='" & Av_Path & "'" &_
				", Av_LinkPopUp='" & Av_LinkPopUp & "'" &_
				" Where av_id=" & av_id
				rs.open sql,con,1
				set rs=nothing
				
		   		set Upload=nothing
		   		response.Write "<script language=""JavaScript"">" & vbNewline &_
					"<!--" & vbNewline &_
					"window.opener.location.reload();" & vbNewline &_
					"window.close();" & vbNewline &_
					"//-->" & vbNewline &_
					"</script>" & vbNewline
				response.end
		End If
	else
		response.write "No User Roles"
	end if 'if session("LstRole")="0ad"
%>