<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	Dim rs
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		EventId=Cint(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if
	
	if Request.QueryString("action")="Update" then
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		
		IsHomeEvent=Upload.Form("IsHomeEvent")
		if IsNumeric(IsHomeEvent) and IsHomeEvent<>"" then
			IsHomeEvent=Cint(IsHomeEvent)
		else
			IsHomeEvent=0
		end if
		IsCatHomeEvent=Upload.Form("IsCatHomeEvent")
		if IsNumeric(IsCatHomeEvent) and IsCatHomeEvent<>"" then
			IsCatHomeEvent=Cint(IsCatHomeEvent)
		else
			IsCatHomeEvent=0
		end if
		EventName=Trim(Replace(Upload.Form("EventName"),"'","''"))
		EventName=Trim(Replace(EventName,"""","&quot;"))
		if EventName="" then
			sEventName="Bắt buộc"
		end if
		
		'EventImages
		set EventImages = Upload.Files("EventImages")
		If EventImages Is Nothing Then
			EventImagesPath=""
		else
		   Filetype = Right(EventImages.Filename,len(EventImages.Filename)-Instr(EventImages.Filename,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"swf" then
				sError=True
				sEventImagesPath=".jpg hoặc .gif"
		   else
		   		EventName_None=Uni2NONE(EventName)
				EventName_None=Replace(EventName_None," ","_")
				EventName_None=Replace(EventName_None,"/","_")
				EventName_None=Replace(EventName_None,"\","_")
		   		EventImagesPath=EventName_None & second(now) & Minute(now) & Hour(now) & Day(now) & Month(now) & Year(now) & "." & Filetype
		   end if
		End If
		
		CatId=Cint(Upload.Form("CatId_DependRole"))
		if CatId=0 then
			sCatId="Bắt buộc"
		end if
		LanguageId=Upload.Form("languageid")
		
		if sCatId="" and sEventName="" then
			'response.write Path & "\" & EventImagesPath
			if EventImagesPath<> "" then
				EventImages.SaveAs Path & "\" & EventImagesPath
			end if
			StatusID=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
			sql="UPDATE Event set"
			sql=sql & " EventName=N'" & EventName & "'"
			if EventImagesPath<> "" then
				sql=sql & ",EventImages='" & EventImagesPath & "'"
			end if
			sql=sql & ",CategoryId=" & CatId
			sql=sql & ",IsHomeEvent=" & IsHomeEvent
			sql=sql & ",IsCatHomeEvent=" & IsCatHomeEvent
			sql=sql & ",statusId='" & StatusID & "'"
			sql=sql & ",Languageid='" & LanguageId & "'"
			sql=sql & ",Approver=N'" & session("user") & "'"
			sql=sql & ",ApproverDate='" & now() & "'"
			sql=sql & " WHERE EventId=" & EventId
			
			set rs=server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			'rs.open sql,con,1 : Quyen Ghi
			'rs.open sql,con,3 : Quyen Doc
			set rs=nothing
			response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
			response.End()
		end if
	else
		sql="SELECT * from event where eventid=" & eventid
		set rs=server.createObject("ADODB.Recordset")
		rs.open sql,con,3
			CatId=Cint(rs("CategoryId"))
			EventName=Trim(rs("Eventname"))
			if rs("IsHomeEvent") then
				IsHomeEvent=1
			else
				IsHomeEvent=0
			end if
			if rs("IsCatHomeEvent") then
				IsCatHomeEvent=1
			else
				IsCatHomeEvent=0
			end if
			languageid=rs("languageid")
		rs.close
		Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fInsertEvent" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.ServerVariables("QUERY_STRING")%>&action=Update" ENCTYPE="multipart/form-data">
  <table width="100%" border="0" cellspacing="2" cellpadding="1">
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Sửa đổi sự kiện</strong></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="checkbox" name="IsHomeEvent" value="1"<%if IsHomeEvent=1 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif"><strong>Trang chủ</strong></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="checkbox" name="IsCatHomeEvent" value="1"<%if IsCatHomeEvent=1 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif"><strong>Trang chuyên 
        mục</strong></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Tên sự kiện:</font></td>
      <td><input name="EventName" type="text" id="EventName" size="35" maxlength="200" value="<%=EventName%>"><font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sEventName%></font>)</font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ảnh minh họa:</font></td>
      <td><input type="file" name="EventImages" id="EventImages" size="21"><font size="1" face="Arial, Helvetica, sans-serif"><font color="#FF0000"><%=sEventImagesPath%></font></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Chuyên mục:</font></td>
      <td>
      	<%Call List_Category_Depend_Role(CatId, "L&#7921;a ch&#7885;n","NONE",Session("LstRole"),"ap",0)%>
      	<font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sCatId%></font>)</font>
      </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ngôn ngữ:</font></td>
      <td><%Call List_Language(languageid)%></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="30" valign="middle"> 
        <input type="submit" name="Submit" value="Sửa đổi">
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();">
      </td>
    </tr>
  </table>
</form>
</body>
</html>