<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/constant.asp"-->
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script type="text/javascript" src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
</head>
<body>
    <div class="container-fluid">
        <%
        Call header()
        %>
    </div>
    <div class="container-fluid">
        <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <table class="table-bordered table">
                <tr>

                    <th colspan="4">QUẢN LÝ CƠ SỞ</th>

                </tr>
                <tr>
                    <th>STT</th>
                    <th>Tên cơ sở</th>
                    <th>Địa chỉ</th>
                    <th><a onclick="Navi_baseEdit('base-add','-1','VN')">Tạo Mới </a></th>
                </tr>
                <%
            sqlL = "SELECT  * FROM  Company  "
            set rsL = Server.CreateObject("ADODB.Recordset")
            rsL.open sqlL,con,1
            IF NOT rsL.EOF THEN
                stt_= 1
                Do while  Not rsL.EOF
                   ' csID_    = rsL("ID")
                   ' csName_    = rsL("company")
                    csTel_     = rsL("Tel")
                    csHotline_ = rsL("Hotline")
                    csAddress_ = rsL("address")
                    csEmail_   = rsL("Email")
                                 
                %>
                <tr>
                    <td><%=stt_ %></td>
                    <td><%=csName_ %></td>
                    <td><%=csAddress_ %></td>
                    <td><a onclick="Navi_baseEdit('Edit','<%=csID_ %>','VN')">Sửa </a>| <a onclick="Navi_baseDel('Del','<%=csID_ %>','VN')">Xóa</a></td>
                </tr>
                <%        
                stt_ = stt_ + 1
                   
                rsL.MoveNext
                Loop
            END IF
                %>
            </table>


            <script type="text/javascript">
                $("#btnsubmit").click(function () {
                    if ($('#company').val() == '') {
                        $('#company').focus();
                        swal("BQT", "Hãy nhập tên cơ sở.");
                    }
                    else {
                        Navi_base('add', '0');
                    }
                });
                function isEmail(email) {
                    var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                    return regex.test(email);
                }
            </script>
        </div>
    </div>


    <%Call Footer()%>
</body>

<script type="text/javascript" src="/administrator/skin/script/sweetalert.min.js"></script>
</html>
