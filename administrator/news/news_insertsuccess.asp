<%@  language="VBSCRIPT" codepage="65001" %>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%CatId=0%>
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
</head>
<body>

    <%
	CategoryLoai	=	GetCategoryLoai(Request.QueryString("CatId"))

	links	="news_addedit.asp"
    %>

    <div class="container-fluid">
        <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <table width="770" border="0" align="center" cellpadding="6" cellspacing="0" style="border: #CCCCCC solid 1;">
                <tr>
                    <td></td>
                    <td>
                        <font size="2" face="Arial, Helvetica, sans-serif"><strong>Cập nhật thành công!!</strong></font>
                        <ul>
                            <li><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; <a href="<%=links%>?iStatus=add"><img src="../../images/icons/icon_key_points.gif" width="48" height="48" border="0" align="absmiddle">Nhập tin mới</a></font>
                                <li><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;<a href="<%=links%>?iStatus=edit&NewsId=<%=Request.Querystring("NewsId")%>&CatId=<%=Request.Querystring("CatId")%>"><img src="../../images/icons/blog-post-edit-icon.jpg" width="48" height="48" border="0" align="absmiddle"> Sửa tin cũ </a></font>
                        </ul>
                    </td>
                </tr>
            </table>

        </div>
    </div>
    <%Call Footer()%>
</body>
</html>
