<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
	

	'Kiem tra du lieu nhap vao
	Dim Upload 'Su dung AspUpload
	Dim sError
	sError="" 'Luu ca'c thong ba'o loi
	Set Upload = Server.CreateObject("Persits.Upload")

	Upload.SetMaxSize 10000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save	
	
	IDAdd=Trim(Upload.Form("IDAdd"))
	bExist	=	false
	sqlCD	=	"Select * from ColorTable where {fn UCASE(ID)}='"& Ucase(IDAdd) &"'"
	Set rsCD = Server.CreateObject("ADODB.Recordset")
	rsCD.open sqlCD,con,1
	if not rsCD.eof then
		bExist	=	true
	else
		bExist	= false
	end if
	set	rsCD = nothing
if bExist = false then	
	set imgfile = Upload.Files("PicColorFile")
	If imgfile Is Nothing Then
		PathColor=""
	else
	   Filetype = Right(imgfile.Filename,len(imgfile.Filename)-Instr(imgfile.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" then
			sError=sError & "&nbsp;-&nbsp; Ảnh màu không phải dạng ảnh<br>"
	   else
			imagename=imgfile.Filename			
			imgfile.SaveAs server.MapPath("/images/ColorTable/")& "\" & imagename
			PathColor	=	"/images/ColorTable/"&imagename
	   end if
	End If	

	
	if  (PathColor <>"" or PathColor <> NULL) and (IDAdd <>"" or IDAdd <> NULL) then
		sqlCD =	"insert into ColorTable(ID,PathColor) values('"& IDAdd &"','"&PathColor&"')"
		Set rsCD = Server.CreateObject("ADODB.Recordset")
		rsCD.open sqlCD,con,1
		set	rsCD = nothing
	end if
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		alert('Đã cập nhật thành công');" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_	
			"	//-->" & vbNewline &_
			"</script>" & vbNewline	
else
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		alert('Mã bạn thêm mới đã bị trùng mới bạn kiểm tra lại');" & vbNewline &_
			"		window.close();" & vbNewline &_	
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
end if
	

	%>

</body>
</html>
