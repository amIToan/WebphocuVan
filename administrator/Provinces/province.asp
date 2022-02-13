<%session.CodePage=65001%>
<%
	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link href="/administrator/font-awesome/css/font-awesome.css" rel="stylesheet" />
</head>
<body>

    <div class="container-fluid">
        <%
            
        Call header()
        %>
    </div>

    <div class="container-fluid">
        <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <%
                idsta_  =  Request.Form("statusId")
    
                IF idsta_ <> ""  And  idsta_ = "add"  THEN 

                    Pcode_ = Request.Form("Pcode")
                    PName_ = Request.Form("PName")
                    IF  PName_ <> ""  THEN
                        sqlp_ = "INSERT INTO [Tb_Provinces]([code],[Name]) VALUES ('"&Pcode_&"',N'"&PName_&"') "
                        on error resume next
                        con.Execute sqlp_
                        if err<>0 then
                            Response.Redirect "/administrator/Provinces/province.asp"
                        else
                          ' Response.Write  "<script>  alert(  swal({title:'',text: 'Lỗi rồi',type: 'error',timer: 2000,showConfirmButton: false})); </script>"
                            Response.Redirect "/administrator/Provinces/province.asp"
                            idsta_ =  ""
                        end if
                    END IF
                ELSEIF idsta_ <> ""  And  idsta_ = "update"  THEN
                    id_ = Request.Form("id_")
                    
                    IF  id_ <> ""  And IsNumeric(id_) THEN                                 
                        Pcode_ = Request.Form("Pcode")
                        PName_ = Request.Form("PName")
                        sql2_ = "UPDATE Tb_Provinces SET code  = '"&Pcode_&"',  Name =N'"&PName_&"' WHERE  id = '"&id_&"'"
                        on error resume next
                            con.Execute sql2_
                            if err<>0 then
                                Response.Redirect "/administrator/Provinces/province.asp?id_="&id_&""
                            else
                                Response.Redirect "/administrator/Provinces/province.asp?id_="&id_&""
                                idsta_ =  ""
                            end if
                    END IF
                ELSE





                    idp_  =  Request.QueryString("id_")
                    IF  idp_ <> "" And IsNumeric(idp_) THEN                   
                    sql1_ =  "SELECT  * FROM  Tb_Provinces WHERE  ID  =  '"&idp_&"'"
                    set rs1_ =  Server.CreateObject("ADODB.Recordset")                
                    rs1_.open sql1_,con,1

                    IF Not rs1_.EOF  THEN
                        name_ = rs1_("Name")
                        code_ = rs1_("code")
                    END IF
            %>
            <form name="f2" id="f2" method="post">
                <div class="col-md-8 col-md-offset-1">
                    <div class="form-group " style="padding-top: 30px;">
                        <label>Tỉnh / Thành phố :</label>
                        <input name="PName" class="form-control" value="<%=name_ %>" placeholder="Tỉnh/ Thành phố..." />
                    </div>
                    <div class="form-group">
                        <label>Mã Tỉnh/TP :</label>
                        <input name="Pcode" class="form-control" value="<%=code_ %>" placeholder="Mã tỉnh" />

                    </div>
                    <div class="form-group">
                        <button type="submit" class="btn btn-primary">Update</button>
                        <input type="hidden" name="statusId" value="update" />
                        <input type="hidden" name="id_" value="<%=idp_ %>" />
                    </div>
                     
                    
                </div>
            </form>
            <%       
                    ELSE  
            %>
            <form name="f1" id="f1" method="post">
                <div class="col-md-8 col-md-offset-1">
                    <div class="form-group " style="padding-top: 30px;">
                        <label>Tỉnh / Thành phố :</label>
                        <input name="PName" id="PName" class="form-control" value="" placeholder="Tỉnh/ Thành phố..." />
                    </div>
                    <div class="form-group">
                        <label>Mã Tỉnh/TP :</label>
                        <input name="Pcode" id="Pcode" class="form-control" value="" placeholder="Mã tỉnh" />

                    </div>
                    <div class="form-group">
                        <button type="submit" class="btn btn-primary">Add</button>
                        <input type="hidden" name="statusId" value="add" />
                    </div>
                </div>
            </form>
            <%     
                    END IF
            %>

            <%    
                End IF
            %>



            <%
                sql_  = "SELECT  * FROM [Tb_Provinces] "
                set rsp_ =  Server.CreateObject("ADODB.Recordset")
                rsp_.open sql_,con,1
                IF Not rsp_.EOF THEN 
            %>

            <div class="col-md-12">

                <%
                    stt_ = 1
                    do while  Not rsp_.EOF
                    Name_ = rsp_("Name")
                    id_   = rsp_("id")
                    IF  Len(stt_) < 2 Then 
                        s= 0&stt_
                    ELSE
                        s =  stt_
                    End  IF

                %>
                <div class="col-md-3">
                    <a href="province.asp?id_=<%=id_ %>"><%=s %>. <%=Name_ %></a>&nbsp;
                <i class="fa fa-trash-o" aria-hidden="true" style="color: #F00;" onclick="Fs_prcDel('prc-del',<%=id_ %>,'0')"></i>
                <a href="hotel.asp?key_=list&id_=<%=id_ %>"><i class="fa fa-list" aria-hidden="true" style="color: #F00;"></i></a>
                </div>
                <% 
                    stt_ = stt_+  1      
                    rsp_.MoveNext
                    Loop
                %>
            </div>
            <%
                End IF
            %>
        </div>
    </div>

    <%Call Footer()%>
    <script src="/administrator/skin/script/sweetalert.min.js"></script>

</body>
</html>
