<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/include/Fs_liblary.asp" -->
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
    <script src="/administrator/js/jquery-2.2.2.min.js"></script>
    <link href="/administrator/css/sweetalert.css" rel="stylesheet" />
    <link href="../../css/styles.css" rel="stylesheet" type="text/css"/>
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <script type="text/javascript" src="../../ckeditor/ckeditor.js"></script>
    <script type="text/javascript" src="../../ckfinder/ckfinder.js"></script>
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>

<body>
<div class="container-fluid">
        <%Call header()%>
</div>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
        <%
            sid =  Request.QueryString("sid")
        id =  Request.QueryString("id")

        IF   sid  =  "add"  THEN
             Set Upload = Server.CreateObject("Persits.Upload")
	         Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	         Upload.codepage=65001
	         Upload.Save

                TxtComany = Upload.Form("TxtComany")
                TxtWeb = Upload.Form("TxtWeb")
                TxtAdress = Upload.Form("TxtAdress")

                set uploadImg = Upload.Files("FileImg")
	            If uploadImg Is Nothing Then
	            	FileImg_=""
	            else
	               Filetype = Right(uploadImg.FileName,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	               	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
	            		sError=True
	            		FileImg_=""
	               else
                       dt = Trim(Replace(getDateServer(),"/",""))
                       dt = Trim(Replace(dt,":",""))
                       dt = Trim(Replace(dt," ",""))             
	            		FileImg_="Pat_"&dt&"."&Filetype
	               end if
	            End If
	            Path=server.MapPath("/images_upload/IMG_Customer")
	            if FileImg_ <>"" then 
	    	        uploadImg.SaveAs Path &"\"&FileImg_
	            end if

		sql	=	"INSERT INTO patner (AvName,AvImg,[Address],Webstite,[view],DateCreate)"
		sql	=	sql+" VALUES (N'"&TxtComany&"', N'"&FileImg_&"', N'"&TxtAdress&"',N'"&TxtWeb&"', '0' , GETDATE() )"
		Set rs=Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1

        Response.Redirect "paner_list.asp"
		set rs = nothing

        ELSEIF   sid  =  "FormAdd"  THEN
            Call  Form_patner(0,0)
        ELSEIF   sid  =  "edit"  and  id <> "" And IsNumeric(id) THEN
            Call  Form_patner(1,id)
        ELSEIF   sid  =  "update"  and  id <> "" And IsNumeric(id) THEN
            
             Set Upload = Server.CreateObject("Persits.Upload")
	         Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	         Upload.codepage=65001
	         Upload.Save

            TxtComany = Upload.Form("TxtComany")
            TxtWeb = Upload.Form("TxtWeb")
            TxtAdress = Upload.Form("TxtAdress")
            txtview= Upload.Form("txtview")
            oldFileImg= Upload.Form("oldFileImg")

            set uploadImg = Upload.Files("FileImg")
	        If uploadImg Is Nothing Then
	        	FileImg_=""
	        else
	           Filetype = Right(uploadImg.FileName,len(uploadImg.Filename)-Instr(uploadImg.Filename,"."))
	           	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
	        		sError=True
	        		FileImg_=""
	           else
                   dt = Trim(Replace(getDateServer(),"/",""))
                   dt = Trim(Replace(dt,":",""))
                   dt = Trim(Replace(dt," ",""))             
	        		FileImg_="Pat_"&dt&"."&Filetype
	           end if
	        End If
	        Path=server.MapPath("/images_upload/IMG_Customer")
	        if FileImg_ <>"" then 
	            uploadImg.SaveAs Path &"\"&FileImg_
            else
                FileImg_=oldFileImg
	        end if

            sql_update="UPDATE patner SET AvName='"&TxtComany&"',AvImg=N'"&FileImg_&"' ,[Address]= N'"&TxtAdress&"',Webstite=N'"&TxtWeb&"',[view]='"&txtview&"' WHERE ID='"&id&"'"
            response.write sql_update
		    Set rs_update=Server.CreateObject("ADODB.Recordset")
		    rs_update.open sql_update,con,1
            set rs_update = nothing
            Response.Redirect "paner_list.asp"

        ELSEIF   sid  =  "del"  and  id <> "" And IsNumeric(id) THEN
            FileName = getColVal("patner","AvImg", " ID = '"&id&"'")        
            sql	=	"DELETE  patner WHERE  ID = '"&id&"'"
            Set rs=Server.CreateObject("ADODB.Recordset")
		    rs.open sql,con,1
            'del file                
            UriF = "/images_upload/IMG_Customer/"&FileName&""                   
            DelFile(UriF)   
            Response.Redirect "paner_list.asp"
        END IF
        %>
  <%       
sub Form_patner(Loai,id)
        
        IF Loai = 1 THEN

            ids = "update"
             sqlp = "Select * From patner  Where Id = '"&id&"'"

                Set rs=Server.CreateObject("ADODB.Recordset")
		        rs.open sqlp,con,1
               IF Not rs.EOF THEN
                id        = Trim(rs("id"))
                Logo_ct        = Trim(rs("AvImg"))
                AvName      = Trim(rs("AvName"))
                diachi     = Trim(rs("Address"))
                DateCreate  = Trim(rs("DateCreate"))
                Webstite    = Trim(rs("Webstite"))
                view        = Trim(rs("view"))
               
                if    Logo_ct <> "" or Not IsNull(Logo_ct) then
                    Logo_ = "<img  src='/images_upload/IMG_Customer/"&Logo_ct&"'"
                else
                    Logo_ = "ko có ảnh"
                end if
                
                if view  = 0  then 
                    view_ = "Không"
                else
                    view_ = "Có"                
                end if
                END IF
        ELSE
            ids = "add"
END IF           
%>      
    <form name="FPalAdd" id="FPalAdd" method="post" action="pal_add.asp?sid=<%=ids %>&id=<%=id %>"" enctype="multipart/form-data">
        <table class=" Tb-input Tb-in w3-table w3-table-all">
            <tr>
                <th colspan="2" class="CTitleClass_AI" style="text-transform: uppercase;">
                    <%if ids = "update" then %>
                        <i class="fa fa-pencil-square-o"></i> CẬP NHẬT ĐỐI TÁC
                    <%else %>
                        <i class="fa fa-plus-square" aria-hidden="true"></i> THÊM ĐỐI TÁC
                    <%end if %>
                </th>
            </tr>
            <tr>
                <td style="width: 20%;">Tên doanh nghiệp:</td>
                <td>
                    <input name="TxtComany" type="text" id="TxtComany" value="<%=AvName %>" />
                </td>
            </tr>
            <tr>
                <td style="width: 20%;">Website:</td>
                <td>
                    <input name="TxtWeb" type="text" id="TxtWeb" value="<%=Webstite %>" />
                </td>
            </tr>
            <tr>
                <td style="width: 20%;">Logo</td>
                <td>
                    <%=Logo_ %>
                    <br />
                    <input name="FileImg" type="file" id="FileImg" value="" />
                    <input name="oldFileImg" type="hidden" id="oldFileImg" value="<%=Logo_ct %>" />
                </td>
            </tr>
            <tr>
                <td style="width: 20%;">Địa chỉ:</td>
                <td>
                    <input name="TxtAdress" type="text" id="TxtAdress" value=" <%=diachi %> " />
                </td>
            </tr>
             <tr>
                <td style="width: 20%;">Hiển thị:</td>
                <td>
                     <input class="w3-check" name="txtview" type="checkbox" id="txtview" value="1" <%if view<>0 then %> checked <%end if %> />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <div>
                         <button class="w3-btn w3-red w3-round" id="PalAdd" name="PalAdd">
                            <%if ids = "update" then %>
                               <i class="fa fa-pencil-square-o"></i> Cập Nhật
                            <%else %>
                               <i class="fa fa-plus-square" aria-hidden="true"></i> Thêm mới
                            <%end if %>
                        </button> 
                    </div>
                </td>

            </tr>
        </table>
    </form>
    <script type="text/javascript">
        $('#PalAdd').click(function () {
            if ($('#TxtComany').val() == '') {
                $('#TxtComany').focus();
                swal("BQT", "Xin vui lòng nhập tên doanh nghiệp.");
            }

            else {
                $('#FPalAdd').submit();
            }
        });
    </script>
    <%end sub %>
    <script src="/interfaces/js/sweetalert.min.js"></script>
    </div><!---/.col-md-10-->
</div><!---/.container-fluid-->
</body>














