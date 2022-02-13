<%@ Language=VBScript CODEPAGE = "65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/aspupload.asp" -->
<%CatId=0%>

<%

 ' Dim uploadsDirVar
 ' uploadsDirVar = path
Dim sError
NewsID =CLng(Request.Querystring("NewsId"))
PictureCaption = Request.Querystring("PictureCaption")
PictureAuthor =Request.Querystring("PictureAuthor")
CatId = CLng(Request.Querystring("CatId"))
PictureAlign=Request.Querystring("PictureAlign")
PictureDirection=CLng(Request.Querystring("PictureDirection"))  
userr=Request.Querystring("userr")
old_PictureId= Request.Querystring("PictureId")
%>


<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<style>
BODY {background-color: white;font-family:arial; font-size:12}
body,td,th {
	font-size: 12px;
}
</style>
<script>
function onSubmitForm(){
    var formDOMObj = document.frmSend
	var img1 = formDOMObj.smallpicturefilename.value
	var img2 = formDOMObj.largepicturefilename.value
	//var imgtype1 = img1.split(".")
	//var imgt1 = imgtype[imgtype.length - 1] 
    if (img1 == "" && img2 == "")
        alert("Chưa chọn file ảnh")
		//return false;
   else if (img1 == "" && img2 != "")
	  alert("Chưa chọn ảnh nhỏ , nếu chọn 1 ảnh thì bắt buộc phải là ảnh nhỏ")
	//  return false;
	//else if (imgt != "jpg" && imgt != "gif" && imgt != "jpeg")
	//   alert(" Ảnh nhỏ không phải là file ảnh ỏ")
    else	
		//var imgtype1 = img1.split(".")
	  //  var imgt1 = imgtype[imgtype.length - 1]    
	   // imgtype = img2.split(".");
	 //   imgt = imgtype[imgtype.length - 1] ;
	 //   if  img2 != "" && (imgt != "jpg" && imgt != "gif" && imgt != "jpeg")
	 //    alert(" Ảnh lớn không phải là file ảnh ỏ")
		 return true;
	  //  else
 return false;
}
</script>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Call header()
	
	Title_This_Page="Tin t&#7913;c -> Nh&#7853;p ảnh"
	
%>
<br><br>
<div style="border-bottom: #A91905 2px solid;font-size:16">Chọn ảnh để đính kèm tin tức</div>

<table width="770" border="0" align="center" cellpadding="6" cellspacing="0">
   <tr><td></td><td>
   	<font size="2" face="Arial, Helvetica, sans-serif"><strong><%=sError%></strong></font>
   	<ul>
   		<li><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;Tiếp tục <a href="news_addnew.asp">nhập tin mới</a></font>
   		<li><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;<a href="news_edit.asp?NewsId=<%=NewsId%>&CatId=<%=CatId%>)">Sửa tin</a> vừa cập nhật</font>
   	</ul>
   	</td></tr>
</table>
  <form name="frmSend" method="POST" enctype="multipart/form-data" action="news_nhapanh.asp">
	<B>File names:</B><br>
    Ảnh Nhỏ: <input name="smallpicturefilename" type="file" size=35><br>
    Ảnh Lớn: <input name="largepicturefilename" type="file" size=35>
    <table width="770" border="0" align="center" cellpadding="6" cellspacing="0">
      <tr>
        <td width="125"> Phải:
            <input name="PictureAlign" type="checkbox" id="PictureAlign" value="<%=PictureAlign%>" checked></td>
        <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">Tác giả ảnh:</font></td>
        <td>          <input name="PictureAuthor" type="text" id="PictureAuthor" value="<%=PictureAuthor%>" size="35" maxlength="50"></td>
      </tr>
      <tr>
        <td>D&#432;ới:
            <input name="PictureDirection" type="checkbox" id="PictureDirection2" value="<%=PictureDirection%>" checked></td>
        <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">Chú thích ảnh:</font></td>
        <td><input name="PictureCaption" type="text" id="PictureCaption" value="<%=PictureCaption%>" size="35" maxlength="200"></td>
      </tr>
    </table>
    <input style="margin-top:4" type=submit value="Upload">
     <input name="NewsId" type="hidden" size=35 value="<%=NewsId%>"><br>
  
     <input name="CatId" type="hidden" size=35 value="<%=CatId%>"><br>
  
     <input name="PictureId" type="hidden" size=35 value="<%=old_PictureId%>"><br>
 
    
</form> 

<%

'OutputForm()
sError = "1"
Dim Upload
Set Upload = New ASPUpload
Upload.Save(path)
NewsID =Upload.Form("NewsId")
PictureCaption = Upload.Form("PictureCaption")
PictureAuthor =Upload.Form("PictureAuthor")
CatId = Upload.Form("CatId")
PictureAlign=Upload.Form("PictureAlign")
PictureDirection=Upload.Form("PictureDirection")  
userr=Upload.Form("userr")
old_PictureId = Upload.Form("PictureId")
dim mt, filekey
		mt = Upload.UploadedFiles.keys
		smallpic=0
		largepic=0
		if (UBound(mt) <> -1) then
	  for each filekey in Upload.UploadedFiles.keys
		     if filekey = "smallpicturefilename" then
		      smallpic = smallpic + 1
			 end if  
			 if filekey="largepicturefilename" then
               largepic = largepic + 1
			 end if 
		  next
		end if
		If smallpic = 0 Then
			SmallPictureFileName=""
		else
		sError = ""
		   smallpicture =replace(Upload.UploadedFiles("smallpicturefilename").FileName,"'","''")
		   Filetype = Right(smallpicture,len(smallpicture)-Instr(smallpicture,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg"then
				sError=sError & "&nbsp;-&nbsp; Ảnh nhỏ không phải là file ảnh<br>"
		 set smt=Server.CreateObject("Scripting.FileSystemObject")
			    smt.DeleteFile Upload.UploadedFiles("smallpicturefilename").Path
		   else
		     if old_PictureId <> 0 then
		      PictureId = old_PictureId
		     else 
		   		PictureId=GetMaxId("Picture", "PictureId", "") 		
		     end if
			 SmallPictureFileName="small_" & PictureId & "." & Filetype
		   end if
		End If
		'largepicture = replace(Upload.UploadedFiles("largepicturefilename").FileName,"'","''") 'request.Files("LargePictureFileName")
		If largepic = 0 Then
			LargePictureFileName=""
		elseif SmallPictureFileName="" then 'Nếu có ảnh lớn nhưng không có ảnh nhỏ
			sError=sError & "&nbsp;-&nbsp; <i>Khi bạn có nhập ảnh:</i> Ảnh nhỏ là bắt buộc, ảnh lớn có thể không có.<br>"
		set smt=Server.CreateObject("Scripting.FileSystemObject")
			    smt.DeleteFile Upload.UploadedFiles("largepicturefilename").Path
		else
		   largepicture = replace(Upload.UploadedFiles("largepicturefilename").FileName,"'","''")
		   Filetype = Right(largepicture,len(largepicture)-Instr(largepicture,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg"then
				'sError=sError & "&nbsp;-&nbsp; Ảnh lớn không phải là file ảnh<br>"
				set smt=Server.CreateObject("Scripting.FileSystemObject")
			    smt.DeleteFile Upload.UploadedFiles("largepicturefilename").Path
		   else
		   		LargePictureFileName="large_" & PictureId & "." & Filetype
		   end if
		End If
		
Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			
			if SmallPictureFileName <>""  and sError = "" then 'Chỉ cần có ảnh nhỏ cũng insert vào mục quản lý ảnh
				'smallpicture.SaveAs Path & "\" & SmallPictureFileName			
			set smt=Server.CreateObject("Scripting.FileSystemObject")
			' smt.MoveFile Path & request.UploadedFiles("smallpicturefilename").FileName, path & SmallPictureFileName
			 If smt.FileExists(path & "\" & SmallPictureFileName) = True Then 
                 smt.DeleteFile path & "\" & SmallPictureFileName, True 
             end if
			smt.MoveFile Upload.UploadedFiles("smallpicturefilename").Path, path & "\" & SmallPictureFileName 
			 set smt=nothing
				if LargePictureFileName<>"" and sError = "" then
				set smt=Server.CreateObject("Scripting.FileSystemObject")
				If smt.FileExists(path & "\" & LargePictureFileName) = True Then 
                 smt.DeleteFile path & "\" & LargePictureFileName, True 
                end if
			    smt.MoveFile Upload.UploadedFiles("largepicturefilename").Path, path & "\" & LargePictureFileName
			    set smt=nothing
									'largepicture.SaveAs Path & "\" & LargePictureFileName
				end if
				sql="insert into Picture (PictureId,PictureCaption,SmallPictureFileName,"
				sql=sql & "LargePictureFileName,PictureAuthor,Creator,CategoryID,StatusID) values "
				sql=sql & "(" & PictureId
				sql=sql & ",N'" & PictureCaption & "'"
				sql=sql & ",'" & SmallPictureFileName & "'"
				sql=sql & ",'" & LargePictureFileName & "'"
				sql=sql & ",N'" & PictureAuthor & "'" ' PictureAuthor & "'"
				sql=sql & ",N'" & session("user") & "'" 'session("user")
				sql=sql & "," & CatId
				sql=sql & ",'ed')"
				'response.write "sqlPicture=" & sql & "<br>"
				rs.open sql,con,1
			else
				PictureId=0
			end if				
			sql="Update News set"
			sql=sql & " PictureID=" & PictureID
			sql=sql & ",PictureAlign='" & PictureAlign & "'"
			sql=sql & ",PictureDirection='" & PictureDirection & "'" 
			sql=sql & " WHERE NewsId='" &  NewsId & "'"
			rs.open sql,con,1
			if sError = "" then
			'response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
		    response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
			else 
			'response.redirect ("news_nhapanh.asp?newsid=" & NewsId & "&catid=" & CatId & "&PictureCaption=" & PictureCaption & "&PictureAuthor=" & PictureAuthor & "&PictureAlign=" & PictureAlign & "&PictureDirection=" & PictureDirection & "&userr=" & session("user") & "&PictureId=" & old_PictureId)
            end if
%>
<%Call Footer()%>
</BODY>
</HTML>
