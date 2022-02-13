<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet/>
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
        <a href="pal_add.asp?sid=FormAdd" class="w3-padding w3-btn w3-blue w3-round w3-margin-top"><i class="fa fa-plus-square" aria-hidden="true"></i> Thêm đối tác</a>
         <%Call Pal_list()%>
    </div><!---/.col-md-10-->
</div><!---/.container-fluid-->
</body>
</html>


  <%
       sub  Pal_list()
           sql	=	"select  * From patner "
		   Set rs=Server.CreateObject("ADODB.Recordset")
		   rs.open sql,con,1
               IF Not rs.EOF THEN
    %>
        <table class="w3-table w3-table-all w3-padding w3-margin-top">
            <tr>
                <td>Logo</td>
                <td>Tên doanh ngiệp</td>
                <td>Địa chỉ Wesite</td>
                <td>Địa chỉ </td>
                <td>Hiển thị</td>
                <td></td>
            </tr>
    <% 
        Do while  Not rs.EOF
            id        = Trim(rs("id"))
            Logo        = Trim(rs("AvImg"))
            AvName      = Trim(rs("AvName"))
            Address     = Trim(rs("Address"))
            DateCreate  = Trim(rs("DateCreate"))
            Webstite    = Trim(rs("Webstite"))
            view        = Trim(rs("view"))
            if    Logo <> "" or Not IsNull(Logo) then
                Logo_ = "<img  src='/images_upload/IMG_Customer/"&Logo&"'"
            else
                Logo_ = "ko có ảnh"
            end if
            
            if view  = 0  then 
                view_ = "Không"
            else
                view_ = "Có"                
            end if
    %>
            <tr>
                <td><%=Logo_ %></td>
                <td><%=AvName %></td>
                <td><%=Webstite %></td>
                <td><%=Address %></td>
                <td><%=view_ %></td>
                <td>
                    <a class="w3-btn w3-blue w3-round" href="pal_add.asp?sid=edit&id=<%=id%>"><font size="2" face="Arial, Helvetica, sans-serif">
                         <i class="fa fa-pencil-square-o" aria-hidden="true"></i> Sửa
                        </font>
                    </a> 
                    <a class="w3-btn w3-red w3-round" href="pal_add.asp?sid=del&id=<%=id%>">
                        <font size="2" face="Arial, Helvetica, sans-serif"><i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Xóa</font>
                    </a></td>
            </tr>
    <% 
        rs.MoveNext
        Loop
    %>

        </table>
    <%                   
               END IF          
		   set rs = nothing
       end sub
    %>