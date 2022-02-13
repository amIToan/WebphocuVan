<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<%f_permission = administrator(false,session("user"),"m_human")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if %>
<html>
<head>

<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<%	
		PathAgency= Server.MapPath("\images_upload\")
		Set Upload = Server.CreateObject("Persits.Upload")
    	'Upload.SetMaxSize 10000000, True 'Dat kich co upload la` 10MB
		Upload.codepage=65001               
    	Upload.Save
		' đẩy ảnh đại diện lên
		set smallpicture = Upload.Files("SmallPictureFileName")
		If smallpicture Is Nothing Then
			SmallPictureFileName=""
		else
		    filename=smallpicture.Filename
        	Filetype=""
        	Call GetFileName(filename,Filetype)
	    	if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg"  and Lcase(Filetype)<>"png" then
			SmallPictureFileName=""
	    	else
				if Trim(Upload.Form("iStatus")) = "edit" then
					SmallPictureFileName="avatar"& Upload.Form("idnhanvien") & "." & Filetype
				else
	   				SmallPictureFileName="avatar"& Upload.Form("iCount") & "." & Filetype
				End If
	    	end if
		End If
   
	if SmallPictureFileName <>"" then 
        smallpicture.SaveAs PathAgency & "\" & SmallPictureFileName    
       
    end if 

		Hoten =Trim(Upload.form("nickname"))
        Email    =Trim(Upload.form("email"))
		phone    =Trim(Upload.form("phone"))
		zalo    = Trim(Upload.form("zalo"))
		ghichu    = Trim(Upload.form("Ghichu"))
		iStatus	=	Trim(Upload.Form("iStatus"))
		idNhanvien= Trim(Upload.Form("idnhanvien"))
		if iStatus = "edit" then
			sql ="update SupportYahoo set idzalo = N'"&zalo&"', hoten= N'"&Hoten&"',mobile='"& phone&"',email=N'"&Email &"',ghichu=N'"& ghichu&"',picturepath= N'"& SmallPictureFileName&"' where id=" & idNhanvien
			set rs=server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			set	rs = nothing
			response.redirect "/administrator/yahoo/Yahoo_list.asp"   
		else
		sql = "Insert into SupportYahoo (idzalo , hoten, mobile, email, ghichu, picturepath) values ( '"&zalo&"', N'"&Hoten&"', '"&phone&"', N'"&Email&"', N'"&ghichu&"', N'"&SmallPictureFileName&"')"
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		set	rs = nothing
		response.redirect "/administrator/yahoo/Yahoo_list.asp"
		end if
	%>
</body>
</html>
<!-- , N'/images_upload/"&SmallPictureFileName&"' -->
<%Function GetFileName(filename,fileType)
    dim tmp
    tmp=Replace(Uni2NONE(filename)," ","-")
    arrdotted=split(tmp,".")
        if UBound(arrdotted)>1 then 'nếu tên chứa nhiều dấu chấm
            File_type=arrdotted(UBound(arrdotted))
        else
            File_type = Right(tmp,len(tmp)-Instr(tmp,"."))
        end if
    filename=tmp
    fileType=File_type
    GetFileName=true
End Function 
%>
