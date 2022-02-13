<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_editor")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		CatId=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if
	
	if request.QueryString("action")="Insert" then
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.codepage=65001
		Upload.SetMaxSize 1000000, True
		'Max size Upload 1MB
		Upload.Save
		
		set smallpicture = Upload.Files("SmallPictureFileName")
		If smallpicture Is Nothing Then
	   		sSmallPictureFileName="&nbsp;Ảnh nhỏ"
		else
			small_Filetype = Right(smallpicture.Filename,len(smallpicture.Filename)-Instr(smallpicture.Filename,"."))
	   		if Lcase(small_Filetype)<>"jpg" and Lcase(small_Filetype)<>"gif" and Lcase(small_Filetype)<>"jpeg" and Lcase(small_Filetype)<>"PNG" then
				sSmallPictureFileName="&nbsp;Sai loại"
			end if
		end If
	
		set largepicture = Upload.Files("LargePictureFileName")
		If largepicture Is Nothing Then
		Else
	   		large_Filetype = Right(largepicture.Filename,len(largepicture.Filename)-Instr(largepicture.Filename,"."))
	   		if Lcase(large_Filetype)<>"jpg" and Lcase(large_Filetype)<>"gif" and Lcase(large_Filetype)<>"jpeg" and Lcase(small_Filetype)<>"PNG" then
				sLargePictureFileName="&nbsp;(<font color=""#FF0000""><strong>*</strong>&nbsp;Sai loại</font>)"
	   		end if
		End If
	
		PictureCaption=Trim(Replace(Upload.Form("PictureCaption"),"'","''"))
		PictureName=Trim(Replace(Upload.Form("PictureName"),"'","''"))
		PictureAuthor=Trim(Replace(Upload.Form("PictureAuthor"),"'","''"))
		Creator=session("user")
		
		'IsHomePicture
		if IsNumeric(Upload.Form("IsHomePicture")) then
			IsHomePicture=1
		else
			IsHomePicture=0
		end if
		'IsCatHomePicture
		if IsNumeric(Upload.Form("IsCatHomePicture")) then
			IsCatHomePicture=1
		else
			IsCatHomePicture=0
		end if
			
		if IsCatHomePicture=1 or IsHomePicture=1 then
		'Co' quyen phu tra'ch hoac admin
			CategoryID=Clng(Upload.Form("CatId_DependRole"))
			if IsCatHomePicture=1 and Categoryid=0 then
				sCategoryid="* Bắt buộc"
			elseif IsCatHomePicture=0 and Categoryid=0 then
			'Case: khong chon dua anh len mang
				CategoryID=Catid
				statusId="ed"
				Approver="NULL"
				ApproverDate="NULL"
			else
			'if IsCatHomePicture=1 and Categoryid<>0
			'Case: chon dua anh len mang
				StatusID=GetRoleOfCat_FromListRole(CategoryId,Session("LstRole"))
				CategoryID=Clng(Upload.Form("CatId_DependRole"))
				Approver="'" & session("user") & "'"
				ApproverDate="'" & now() & "'"
			end if
		else
			CategoryId=CatId
			IsHomePicture=0
			IsCatHomePicture=0
			StatusID="ed"
			Approver="NULL"
			ApproverDate="NULL"
		end if
		
		if sCategoryid="" and sLargePictureFileName="" and sSmallPictureFileName="" then
			PicId=GetMaxId("Picture", "PictureId", "")
			SmallPictureFileName="small_" & PicId & "." & small_FileType
			smallpicture.SaveAs Path & "\" & SmallPictureFileName
			
			If Largepicture Is Nothing Then
				LargePictureFileName=""
			else
				LargePictureFileName="large_" & PicId & "." & large_FileType
				Largepicture.SaveAs Path & "\" & LargePictureFileName
			end if
			
			sql="insert into Picture (PictureId,PictureName,Picturecaption,smallPictureFilename,largePictureFilename,"
			sql=sql & "PictureAuthor,Creator,CategoryID,StatusID,IsHomePicture,"
			sql = sql & "IsCatHomePicture,Approver,ApproverDate) values "
			sql=sql & "(" & PicId
			sql=sql & ",N'" & PictureName & "'"
			sql=sql & ",N'" & Picturecaption & "'"
			sql=sql & ",'" & smallPictureFilename & "'"
			sql=sql & ",'" & largePictureFilename & "'"
			sql=sql & ",N'" & PictureAuthor & "'"
			sql=sql & ",N'" & Creator & "'"
			sql=sql & "," & CategoryID
			sql=sql & ",'" & StatusID & "'"
			sql=sql & "," & IsHomePicture
			sql=sql & "," & IsCatHomePicture
			sql=sql & "," & Approver
			sql=sql & "," & ApproverDate & ")"
			
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			set rs=nothing
			set Upload=nothing
			response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
			response.End()
		end if 'Of if sCategoryid="" and sLargePictureFileName="" and sSmallPictureFileName="" then
		set Upload=nothing
	end if 'Of request.QueryString("action")="Insert"
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?param=<%=CatId%>&action=Insert" method="post" enctype="multipart/form-data" name="fNew">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr align="center" valign="middle"> 
      <td height="40" colspan="2" valign="middle"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Tạo 
        ảnh mới</strong></font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Tên ảnh: </font></td>
      <td align="left"><input name="PictureName" type="text" id="PictureName" size="33" maxlength="100" value="<%=PictureName%>"></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Chuyên 
        mục:</font></td>
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=GetNameOfCategory(CatId)%></b></font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Ảnh 
        nhỏ:</font></td>
      <td align="left"><input name="SmallPictureFileName" type="file" id="SmallPictureFileName"> <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sSmallPictureFileName%></font>)</font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Ảnh 
        lớn:</font></td>
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="LargePictureFileName" type="file" id="LargePictureFileName"><font size="1" face="Arial, Helvetica, sans-serif"><%=sLargePictureFileName%></font>
        </font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Tác 
        giả:</font></td>
      <td align="left"><input name="PictureAuthor" type="text" id="PictureAuthor" value="<%=PictureAuthor%>" size="33" maxlength="200"></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Chú thích: </font></td>
      <td align="left"><input name="PictureCaption" type="text" id="PictureCaption" size="33" maxlength="500" value="<%=PictureCaption%>"></td>
    </tr>
	<%strRole=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
	if strRole="ap" or strRole="ad" then%>
    <tr> 
      <td colspan="2" align="left"><strong><font size="2" face="Arial, Helvetica, sans-serif"> 
        Đánh dấu vào dưới đây để đưa ảnh vào phần tin ảnh</font></strong></td>
    </tr>
    <tr> 
      <td colspan="2" align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
        &nbsp;&nbsp;&nbsp; 
        <input name="IsHomePicture" type="checkbox" id="IsHomePicture" value="1"<%if IsHomePicture=1 then%> checked<%end if%>>
        Của trang chủ</font><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
    </tr>
    <tr> 
      <td colspan="2" align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
        &nbsp;&nbsp;&nbsp; 
        <input name="IsCatHomePicture" type="checkbox" id="IsCatHomePicture" value="1"<%if IsCatHomePicture=1 then%> checked<%end if%>>
        Của 
        <%Call List_Category_Depend_Role(CategoryId, "Chuy&#234;n m&#7909;c","NONE",Session("LstRole"),"ap",0,0)%>
        </font><font size="1" face="Arial, Helvetica, sans-serif" color="#FF0000"><%=sCategoryId%></font></td>
    </tr>
	<%End if%>
    <tr> 
      <td align="center" colspan="2" height="35" valign="bottom">
	  	<input type="hidden" name="action" value="Insert"> <input type="submit" name="Submit" value="Tạo mới"> 
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();"> 
      </td>
    </tr>
  </table>
</form>
</body>
</html>