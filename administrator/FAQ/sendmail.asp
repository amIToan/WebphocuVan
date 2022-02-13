<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	idYKien = Clng(Request.QueryString("param"))
	action  = Request.QueryString("action")
Select case action 
    case "send"
		sqlyk = "SELECT   tieude, hovaten,email,faq  from Y_KIEN where id='"& idYKien &"'"
		Set rsYKien=Server.CreateObject("ADODB.Recordset")
		rsYKien.open sqlyk,con,1
		if not rsYKien.eof then
			Email=rsYKien("email")
			hovaten=Trim(rsYKien("hovaten"))
			tieude=	Trim(rsYKien("tieude"))
		end if
		set rsYKien = nothing

        CASE "SENDMAIL"

              txtName       =   Request.Form("txtName")
              txtEmail      =   Request.Form("txtEmail")
              txtPass       =   Request.Form("txtPass")
              txtMailTo     =   Request.Form("txtMailTo")
              txtTitle      =   Request.Form("txtTitle")
              txtBody       =   Request.Form("txtBody")

              Call SendEmail(txtBody,txtMailTo, txtTitle,txtEmail,txtPass)

               response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
			"window.location=""sendmail.asp"";" & vbNewline &_
			"alert('Gửi thành công');"& vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline
	    response.End()
        END SELECT
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/jquery.js"></script>
    <script type="text/javascript" src="../../ckeditor/ckeditor.js"></script>
    <script type="text/javascript" src="../../ckfinder/ckfinder.js"></script>
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>

<body>
    <form name="form1" id="form1" method="post" action="sendmail.asp?id=<%=idYKien%>&action=SENDMAIL">
        <table  class="w3-table w3-table-all">
            <tr>
                <th colspan="2" class="CTitleClass_AI">
                    <br>
                    Soạn Mail<br>
                </th>
            </tr>
            <tr>
                <td>Tài khoản (email): </td>
                <td>
                    <input name="txtEmail" type="text" id="txtEmail" value="" size="25" class="w3-input w3-border w3-round" /></td>
            </tr>
            <tr>
                <td>Mật khẩu(email): </td>
                <td>
                    <input name="txtPass" type="text" id="txtPass" value="" size="25" class="w3-input w3-border w3-round" /></td>
            </tr>
            <tr>
                <td>Đến Email </td>
                <td>
                    <input name="txtMailTo" type="text" id="txtMailTo" value="<%=Email%>" size="25" class="w3-input w3-border w3-round" /></td>
            </tr>
            <tr>
                <td>Tiêu đề</td>
                <td>
                    <input name="txtTitle" type="text" id="txtTitle" value="" size="25" class="w3-input w3-border w3-round" /></td>
            </tr>
            <tr>
                <td>Thông điệp  : </td>
                <td>
                    <textarea name="txtBody" cols="50" rows="10" id="txtAns"></textarea>
                </td>
            </tr>

            <tr>
                <td>&nbsp&nbsp&nbsp</td>
                <td align="right">
                    <button class="w3-btn w3-red w3-round" id="btn_send"  name="btn_send" type="submit" name="Submit"><i class="fa fa-paper-plane-o" aria-hidden="true"></i> Gửi đi</button>
                </td>
            </tr>
        </table>
    </form>
        <script type="text/javascript">
            CKEDITOR.replace('txtAns');
    </script>
</body>
</html>
<script type="text/javascript">
    $(document).ready(function () {
        $("#btn_send").click(function () {
            if ($('#txtEmail').val() == '') {
                $('#txtEmail').focus();
                swal("BQT", "Xin vui lòng nhập email.");
            }
            else if (!isEmail($('#txtEmail').val())) {
                $('#txtEmail').focus();
                swal("BQT", "Sai định dạng email.vd: abc@gmail.com");
            }
            else if ($('#txtPass').val() == '') {
                $('#txtPass').focus();
                swal("BQT", "Xin vui lòng nhập password.");
            }
            else if ($('#txtTitle').val() == '') {
                $('#txtTitle').focus();
                swal("BQT", "Vui lòng nhập tiêu đề.");
            } 
            else {
                $("#form1").submit();
            }

        });
        function isEmail(email) {
            var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
            return regex.test(email);
        }
    });
    </script>
<%
sub SendEmail(content,email, Subject,email_send,pass_send)
  mailBody  = content
  '19 - 06 - 2015
   Dim ObjSendMail
   Set ObjSendMail =server.createobject("CDO.Message")
   
   'This section provides the configuration information for the remote SMTP server.
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2'Send the message using the network (SMTP over the network).
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
   
   ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") =email_send 'your Google apps mailbox address
   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") =pass_send 'Google apps password for that mailbox
   
   ObjSendMail.Configuration.Fields.Update
   ObjSendMail.To =Email
  ' ObjSendMail.CC = "<sp@xseo.vn>"
   ObjSendMail.Subject =Subject    '"Phan mem XSEO thong tin dang ky" 
   ObjSendMail.From =email_send 'thay doi tieu de info thi sua thanh hethong<info@xseo.vn>
  '  ObjSendMail.HTML
   'ObjSendMail.IsHTML = True 
    ObjSendMail.BodyPart.Charset = "utf-8" 
   'ObjSendMail.HTMLBodyPart.Charset = "utf-8"
   ' we are sending a text email.. simply switch the comments around to send an html email instead
   ObjSendMail.HTMLBody = mailBody
   
   
   ObjSendMail.Send
   Set ObjSendMail = Nothing 
End sub
%>