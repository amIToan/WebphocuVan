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
                key_ =  Request.QueryString("key_")
                IF key_  = "list" THEN
                    Call Hotel_list()
                END IF                
            %>
        </div>
    </div>

    <%Call Footer()%>
    <script src="/administrator/skin/script/sweetalert.min.js"></script>

</body>
</html>
<%
    sub Hotel_list()        
    sql1_  = " SELECT  n.Title,n.NewsID FROM  News as n  INNER JOIN NewsDistribution as d on  n.newsID =  d.newsID INNER  JOIN  NewsCategory as c on  d.CategoryID = c.CategoryID"
    sql1_ =  sql1_&" WHERE  c.CategoryLoai = 6 "
    set rsh_ =  Server.CreateObject("ADODB.Recordset")
    rsh_.open sql1_,con,1
    IF Not rsh_.EOF THEN 


%>

<div class="col-md-12">
<div class="col-md-12">
    <h4>Danh sách khách sạn : </h4>
</div>
    <%
        stt_ = 1
        do while  Not rsh_.EOF
            Name_ = rsh_("Title")
            id_   = rsh_("NewsID")
            IF  Len(stt_) < 2 Then 
                s= 0&stt_
            ELSE
                s =  stt_
            End  IF
    %>
    <div class="col-md-3">
        <a href="province.asp?id_=<%=id_ %>"><%=s %>. <%=Name_ %></a>&nbsp;          
    </div>
    <% 
        stt_ = stt_+  1      
        rsh_.MoveNext
        Loop
    %>
</div>
<%
    END IF
    end sub
%>