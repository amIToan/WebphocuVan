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
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />

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
            <table class="table table-bordered">
                <tr>
                    <th colspan="2">CHÀO MỪNG ĐẾN VỚI TRANG QUẢN TRỊ</th>
                </tr>
                <tr>
                    <td>NGÀY THÔNG BÁO</td>
                    <td>TÓM TẮT	</td>
                </tr>


                <tr>
                    <td class="text-center">Xếp thứ hạng:</td>
                    <td>
                        <script type="text/javascript" src="http://xslt.alexa.com/site_stats/js/t/a?url=<%=Website_sys%>"></script>
                    </td>
                </tr>
            </table>
        </div>
    </div>


    <!-- jQuery -->
    <script src="/administrator/js2/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="/administrator/js2/bootstrap.min.js"></script>

    <!-- Morris Charts JavaScript -->
    <script src="/administrator/js2/plugins/morris/raphael.min.js"></script>
    <script src="/administrator/js2/plugins/morris/morris.min.js"></script>
    <script src="/administrator/js2/plugins/morris/morris-data.js"></script>

    <%Call Footer()%>
</body>
</html>
