<%@  language="VBSCRIPT" codepage="65001" %>
<%Call PhanQuyen("QLyHeThong")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<script src="/ckeditor/ckeditor.js"></script>
<script src="/ckfinder/ckfinder.js"></script>
<%
    lang = Session("Language")
    if lang = "" then lang = "VN"
if Trim(Request.Querystring("action"))="Update" then
	sError=False
	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save
	
	set uploadImg = Upload.Files("icon")
	If uploadImg Is Nothing Then
		icon=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   if  Lcase(Filetype)<>"ico" then
			sError=True
			icon=""
	   else
			icon="icon." & Filetype
	   end if
	End If
	Path=server.MapPath("/images/logo/")
	if icon <>"" then 
		uploadImg.SaveAs Path & "\" &icon
    'Response.Write icon
	end if

    	set uploadImg = Upload.Files("logoF")
	If uploadImg Is Nothing Then
		logoF=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
			sError=True
			logoF=""
	   else
			logoF="logoF." & Filetype
	   end if
	End If
	Path=server.MapPath("/images/logo/")
	if logoF <>"" then 
		uploadImg.SaveAs Path & "\" &logoF
    'Response.Write logoF
	end if
	
   ' Response.End

	set uploadImg = Upload.Files("logo")
	If uploadImg Is Nothing Then
		logo=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
			sError=True
			logo=""
	   else
			logo="logo." & Filetype
	   end if
	End If
	Path=server.MapPath("/images/logo/")
	if logo <>"" then uploadImg.SaveAs Path & "\" &logo
		
	set uploadImg = Upload.Files("background")
	If uploadImg Is Nothing Then
		background=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
			sError=True
			background=""
	   else
			background=uploadImg.Filename
	   end if
	End If
	Path=server.MapPath("/images/")
	if background <>"" then uploadImg.SaveAs Path & "\" &background
	
	set uploadImg = Upload.Files("banner")
	If uploadImg Is Nothing Then
		banner=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
			sError=True
			banner=""
	   else
			banner=uploadImg.Filename
	   end if
	End If
	Path=server.MapPath("/images/")
	if banner <>"" then uploadImg.SaveAs Path & "\" &banner	

	set uploadImg = Upload.Files("footer")
	If uploadImg Is Nothing Then
		footer1=""
	else
	   Filetype = Right(uploadImg.Filename,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
			sError=True
			footer1=""
	   else
			footer1=uploadImg.Filename
	   end if
	End If
	Path=server.MapPath("/images/")
	if footer1 <>"" then uploadImg.SaveAs Path & "\" &footer1		

	company		=	Trim(replace(Upload.form("company"),"'","''"))
	Tel			=	Trim(replace(Upload.form("Tel"),"'","''"))
	Hotline			=	Trim(replace(Upload.form("Hotline"),"'","''"))
	Email		=	Trim(replace(Upload.form("Email"),"'","''"))
	Website		=	Trim(replace(Upload.form("Website"),"'","''"))
	address		=	Trim(replace(Upload.form("address"),"'","''"))
	calltime	=	Trim(replace(Upload.form("calltime"),"'","''"))
	Masothue	=	trim(replace(Upload.form("Masothue"),"'","''"))
	GPKD		=	trim(replace(Upload.form("GPKD"),"'","''"))
	page_title	=	trim(replace(Upload.form("page_title"),"'","''"))
	InfoOrder	=	trim(replace(Upload.form("infoOrder"),"'","''"))
	meta_description	=	trim(replace(Upload.form("meta_description"),"'","''"))
	meta_keywords		=	trim(replace(Upload.form("meta_keywords"),"'","''"))
	Cfont				=	trim(replace(Upload.form("Cfont"),"'","''"))
	'fsize				=	Cint(Upload.form("fsize"))
	idgoogle			=	replace(Upload.form("idgoogle"),"'","''")
    idfacebook			=	replace(Upload.form("idfacebook"),"'","''")
    idgplus				=	replace(Upload.form("idgplus"),"'","''")
    idyoutube			=	replace(Upload.form("idyoutube"),"'","''")
    idskype				=	replace(Upload.form("idskype"),"'","''")
	introduction		=	replace(Upload.form("introduction"),"'","''")

    idEmail				=	replace(Upload.form("idEmail"),"'","''")
    idvideo				=	replace(Upload.form("idvideo"),"'","''")
    idFax				=	replace(Upload.form("Fax"),"'","''")
    TitleF              =	replace(Upload.form("TitleF"),"'","''")
	
	show_intro_home = Upload.Form("show_intro_home")
	
	if not IsNumeric(Upload.Form("show_intro_home")) then
		show_intro_home=0
	else
		show_intro_home=Clng((Upload.Form("show_intro_home")))
	end if	

    embed_head      =	replace(Upload.form("embed_head"),"'","''")
    embed_footer    =	replace(Upload.form("embed_footer"),"'","''")
	
	Set rs = Server.CreateObject("ADODB.Recordset")	
	sql	=	"Update Company set "
	sql	=	sql + "company = N'"& company &"'"
	sql	=	sql + ",Tel = N'"& Tel &"'"
	sql	=	sql + ",Hotline = N'"& Hotline &"'"
    sql	=	sql + ",Email = N'"& Email &"'"
	sql	=	sql + ",Website = N'"& Website &"'"
	sql	=	sql + ",address = N'"& address &"'"
	sql	=	sql + ",calltime = N'"& calltime &"'"
	sql	=	sql + ",Masothue = N'"& Masothue &"'"
	sql	=	sql + ",GPKD = N'"& GPKD &"'"
	sql	=	sql + ",page_title = N'"& page_title &"'"
	sql	=	sql + ",TitleF = N'"& TitleF &"'"
	sql	=	sql + ",meta_description = N'"& meta_description &"'"
	sql	=	sql + ",meta_keywords = N'"& meta_keywords &"'"
	sql	=	sql + ",Cfont = N'"& Cfont &"'"
	sql	=	sql + ",fsize = '"& fsize &"'"
    sql	=	sql + ",idgoogle = '"& idgoogle &"'"
	sql	=	sql + ",idfacebook = N'"& idfacebook &"'"
	sql	=	sql + ",idgplus	= N'"& idgplus	 &"'"
	sql	=	sql + ",idyoutube= N'"& idyoutube &"'"
	sql	=	sql + ",idskype	= N'"& idskype	 &"'"

    sql	=	sql + ",idEmail	= N'"& idEmail	 &"'"
    sql	=	sql + ",idvideo	= N'"& idvideo	 &"'"
    sql	=	sql + ",Fax	= N'"& idFax	 &"'"

	sql	=	sql + ",introduction = N'"& introduction &"'"
	sql	=	sql + ",show_intro_home = "&show_intro_home
if icon <> "" then	
	sql	=	sql + ",icon = '"& icon &"'"
end if	
if logoF <> "" then	
	sql	=	sql + ",logoF = '"& logoF &"'"
end if	
if logo <> "" then
	sql	=	sql + ",logo = '"& logo &"'"
end if
if background <> "" then	
	sql	=	sql + ",background = '"& background &"'"
end if
if banner <> "" then	
	sql	=	sql + ",banner = '"& banner &"'"
end if
if footer1 <> "" then	
	sql	=	sql + ",footer = '"& footer1 &"'"
end if

    if embed_head <> "" then	
	    sql	=	sql + ",embed_head = '"& embed_head &"'"
    end if

    if embed_footer <> "" then	
	    sql	=	sql + ",embed_footer = '"& embed_footer &"'"
    end if

    sql = sql + " where lang = '"&lang&"'"
	rs.open sql,Con,3	
	response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
			"window.location=""company_edit.asp"";" & vbNewline &_
			"alert('Cập nhận thành công');"& vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline
	response.End()
end if

		
Set rs=Server.CreateObject("ADODB.Recordset")		
sql="select top 1 * from Company where lang = '"&lang&"'"
rs.open sql,con,1
if not rs.eof then
	company		=	Trim(rs("company"))
	Tel			=	Trim(rs("Tel"))
	Hotline			=	Trim(rs("Hotline"))
	Email		=	Trim(rs("Email"))
	Website		=	Trim(rs("Website"))
	address		=	Trim(rs("address"))
	calltime	=	Trim(rs("calltime"))
	Masothue	=	trim(rs("Masothue"))
	GPKD		=	trim(rs("GPKD"))
    Fax         =   trim(rs("Fax"))
	page_title	=	trim(rs("page_title"))
	TitleF	    =	trim(rs("TitleF"))
	meta_description	=	trim(rs("meta_description"))
	meta_keywords		=	trim(rs("meta_keywords"))
	icon				=	trim(rs("icon"))
	logoF				=	trim(rs("logoF"))
	Logo				=	trim(rs("Logo"))
	background			=	trim(rs("background"))
	banner				=	trim(rs("banner"))
	footer1				=	trim(rs("footer"))
	Cfont				=	trim(rs("Cfont"))
	fsize				=	rs("fsize")
	idgoogle			=	rs("idgoogle")
    idfacebook			=	rs("idfacebook")
    idgplus			    =	rs("idgplus")
    idyoutube			=	rs("idyoutube")
    idskype			    =	rs("idskype")
	introduction		=	rs("introduction")
    idEmail             =   rs("idEmail")  
    idvideo			    =	rs("idvideo")
	show_intro_home		=	rs("show_intro_home")
    embed_head		    =	rs("embed_head")
    embed_footer		=	rs("embed_footer")
end if
rs.close
set rs=nothing
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/js2/bootstrap.min.js"></script>  
    <script type="text/javascript" src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link href="../../administrator1/font-awesome/css/font-awesome.css" rel="stylesheet" />
</head>
<body>
 <div class="container-fluid">
        <%Call header()%>
</div>

<div class="container-fluid">
    <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
        <form name="fcompany" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=Update&<%=Request.ServerVariables("QUERY_STRING")%>" enctype="multipart/form-data">
    <div style="width: 1000px; margin: auto; float: left; margin: 10px 20px;">
        <table border="0" align="center" cellpadding="2" cellspacing="2" class="w3-table w3-table-all ">
            <tr align="center" valign="middle">
                <td height="40" colspan="2" valign="middle">
                    <div align="center"><strong>SỬA THÔNG TIN CÔNG TY </strong></div>
                </td>
            </tr>    
            <tr>
                <td width="125">Tên công ty:</td>
                <td width="623">
                    <input name="company" type="text" id="company" class="w3-input w3-border w3-round" value="<%=company%>" size="45" maxlength="100"></td>
            </tr>
            <tr>
                <td>ico</td>
                <td>
                    <%if icon<>"" then%>
                    <img src="/images/logo/<%=icon%>" border="0" style="max-width: 100px; "><br />&nbsp;
                    <%end if%>
                    <input name="icon" type="file" id="icon" >
                    bắt buộc file *.ico </td>
            </tr>
            <tr>
                <td>Logo</td>
                <td>
                    <%if logo<>"" then%>
                    <img src="/images/logo/<%=logo%>" border="0" style="max-width: 100px;"><br>
                    <%end if%>
                    <input name="Logo" type="file" id="logo" size="25">
                    *.png, *.jpeg, *.gif, *.jpg </td>
            </tr>
            <tr>
                <td>
                    <div>logo footer</div>
                </td>
                <td>
                    <%if logoF<>"" then%>
                    <img src="/images/logo/<%=logoF%>" border="0" style="max-width: 100px;"><br>
                    <%end if%>
                    <input name="logoF" type="file" id="logoF" size="25">
                    *.png, *.jpeg, *.gif, *.jpg </td>
            </tr>
            <tr>
                <td>
                    <div>Địa chỉ:</div>
                </td>
                <td>
                    <input name="address" type="text" class="w3-input w3-border w3-round" size="45" maxlength="500" value="<%=address%>">
                </td>
            </tr>
            <tr>
                <td>
                    <div>Điện thoại:</div>
                </td>
                <td>
                    <input name="Tel" type="text" id="Tel" class="w3-input w3-border w3-round" size="45" maxlength="100" value="<%=Tel%>"></td>
            </tr>
            <tr>
                <td>
                    <div>Hotline:</div>
                </td>
                <td>
                    <input name="Hotline" type="text" id="Hotline" class="w3-input w3-border w3-round" size="45" maxlength="100" value="<%=Hotline%>"></td>
            </tr>
                <tr>
                <td>
                    <div>Tên Ngân Hàng:</div>
                </td>
                <td>
                    <input name="Fax" type="text" id="Fax" class="w3-input w3-border w3-round" size="45" maxlength="100" value="<%=Fax %>"></td>
             </tr>
            <tr>
                <td>
                    <div>Email: </div>
                </td>
                <td>
                    <input name="Email" type="text" id="Email" class="w3-input w3-border w3-round" size="45" maxlength="100" value="<%=Email%>">
                </td>
            </tr>
            <tr>
                <td>
                    <div>Website:</div>
                </td>
                <td>
                    <input name="Website" type="text" id="Website" class="w3-input w3-border w3-round" size="45" maxlength="100" value="<%=Website%>"></td>
            </tr>
            <tr>
                <td>
                    <div>Số tài khoản: </div>
                </td>
                <td>
                    <input name="masothue" type="text" size="45" class="w3-input w3-border w3-round" maxlength="100" value="<%=masothue%>"></td>
            </tr>        
            <tr>
                <td>
                    <div>Tên Tài Khoản:</div>
                </td>
                <td>
                    <input name="GPKD" type="text" size="45" class="w3-input w3-border w3-round" maxlength="100" value="<%=GPKD%>"></td>
            </tr>            
            <tr>
                <td>
                    <div>Page title: </div>
                </td>
                <td>
                    <input name="page_title" type="text" size="45" class="w3-input w3-border w3-round" maxlength="500" value="<%=page_title%>"></td>
            </tr>
            <tr>
                <td>
                    <div>Meta description:</div>
                </td>
                <td>
                    <input name="meta_description" type="text" size="45" class="w3-input w3-border w3-round" maxlength="500" value="<%=meta_description%>"></td>
            </tr>
            <tr>
                <td>
                    <div>meta keywords:</div>
                </td>
                <td>
                    <input name="meta_keywords" type="text" size="45" class="w3-input w3-border w3-round" maxlength="500" value="<%=meta_keywords%>"></td>
            </tr>
            <tr>
                <td>
                    <div>Thời gian làm việc: </div>
                </td>
                <td>
                    <input name="calltime" type="text" size="45" class="w3-input w3-border w3-round" maxlength="100" value="<%=calltime%>"></td>
            </tr>
            <tr>
                <td>
                    <div>Mã google cung cấp: </div>
                </td>
                <td>
                    <input name="idgoogle" type="text" size="45" class="w3-input w3-border w3-round" maxlength="100" value="<%=idgoogle%>"></td>
            </tr>
            <tr>
                <td>link fanpage facebook</td>
                <td>
                    <input name="idfacebook" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idfacebook%>"></td>
            </tr>
            <tr>
                <td>link youtube</td>
                <td>
                    <input name="idyoutube" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idyoutube%>"></td>
            </tr>
            <tr>
                <td>id skype</td>
                <td>
                    <input name="idskype" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idskype%>"></td>
            </tr>
            <tr>
                <td>link G+</td>
                <td>
                    <input name="idgplus" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idgplus%>"></td>
            </tr>
            <tr>
                    <td>link Email</td>
                    <td>
                        <input name="idEmail" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idEmail %>"></td>
            </tr>
            <tr>
                    <td>link Video</td>
                    <td>
                        <input name="idvideo" type="text" size="45" class="w3-input w3-border w3-round" value="<%=idvideo %>"></td>
                </tr>
            <tr>
             <tr>
                    <td>TitleF</td>
                    <td>
                        <input name="TitleF" id="TitleF" type="text" size="45" class="w3-input w3-border w3-round" value="<%=TitleF %>"></td>
                </tr>
            <tr>
            <tr>
                    <td>Nhúng header</td>
                    <td>
                        <textarea id="embed_head" name="embed_head" class="w3-input w3-border w3-round" rows="3"><%=embed_head %></textarea>
                    </td>
            </tr>
            <tr>
                    <td>Nhúng footer</td>
                    <td>
                        <textarea id="embed_footer" name="embed_footer" class="w3-input w3-border w3-round" rows="3"><%=embed_footer %></textarea>
                    </td>
            </tr>
            <tr>
                    <td>introduction</td>
                    <td>
                        <textarea id="introduction" name="introduction" class="w3-input w3-border w3-round" style="height:136px;"><%=introduction %></textarea>
                    </td>
            </tr>
            <tr>
                <td colspan="2" class="w3-center">
                    <button class="w3-btn w3-red w3-round" type="submit" name="Submit"><i class="fa fa-pencil-square-o" aria-hidden="true"></i> Cập nhật</button>
                </td>
            </tr>
    </div>
</form>                
    </div>
</div><!---/.container-fluid--->
    <%Call Footer()%>
    <script type="text/javascript">
        CKEDITOR.replace('introduction');
        var editor = CKEDITOR.replace('introduction');
        CKFinder.setupCKEditor(editor, '/ckfinder/');
    </script>
</body>
</html>
