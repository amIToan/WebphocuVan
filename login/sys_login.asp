<%@  language="VBSCRIPT" codepage="65001" %>
<!DOCTYPE html>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp" -->
<!--#include virtual="/include/func_common.asp" -->
<html>
<head>
    <!-- Basic Page Needs -->
    <meta charset="utf-8">
    <title>ART BOX</title>
    <meta name="description" content="">
    <meta name="keywords" content="">
    <meta name="author" content="">



    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <link rel="stylesheet" type="text/css" href="/stylesheets/bootstrap.css">
    <link rel="stylesheet" type="text/css" href="/stylesheets/style.css">
    <link rel="stylesheet" type="text/css" href="/stylesheets/animate.css">
    <link rel="stylesheet" type="text/css" href="/stylesheets/bootstraps.css" />
    <link rel="stylesheet" type="text/css" href="/login/Login.css" />
    <link rel="stylesheet" type="text/css" href="/stylesheets/sweetalert.css" />

    <%'Call Fs_Library_css() %>
</head>
<body class="one-page full-screen bg-login">
    <div class="containner" style="width: 45%; margin: auto;">
        <form name="Flogin_adm" id="Flogin_adm" method="post" onsubmit="Login_system('Login_system');">
            <div class="modal-dialog">
                <div class="modal-body bg-login-w1">
                    <h1 class="login-heading ">WELCOME MANAGER. PLEASE LOGIN...</h1>
                    <input type="text" name="LoginID" placeholder="Account" required="required" class="input-txt">
                    <input type="password" name="PassID" placeholder="Password" required="required" class="input-txt">
                </div>
                <div style="width: 100%; float: left;" class="bg-login-w1">
                    <div class="col-md-9" style="padding-top: 1em;">
                        <span style="color: #ccc;"></span>
                        &nbsp;
                    </div>
                    <div class="col-md-3 bg-login-w1">
                        <button type="submit" class="btn-login">Login</button>
                        &nbsp;
                    </div>
                </div>

            </div>
        </form>
    </div>
    <script src="/javascript/sweetalert.min.js"></script>
    <script src="/javascript/jquery.min.js"></script>
    <script src="/login/Login.js"></script>
</body>
</html>
