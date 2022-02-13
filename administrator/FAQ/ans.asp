<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_faq")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	lang=Request.QueryString("param")
	lang=replace(lang,"'","''")
	action  = Request.QueryString("action")
	valAction = ""
Select case action 					
	case "add"
		Email   = Trim(Request.Form("txtEmail"))
		hovaten = Trim(Request.Form("txtName"))
		tieude  = Trim(Request.Form("txtTitle"))
		noidung = Trim(Request.Form("txtAns"))
		isshow  = Request.Form("isShow")
		if isshow <> "" then
			isshow = 1
		else
			isshow = 0
		end if 
       
		Traloi	=	Trim(Request.Form("txtQuestion"))
		sqlyk	=	"insert into y_kien(hovaten,Email,tieude,noidung,faq,show,Traloi)"
		sqlyk	=	sqlyk+" values( N'"& hovaten &"','"& Email &"',N'"& tieude &"',N'"& noidung &"','1','"& isshow &"',N'"& Traloi &"')"
		Set rsYKien=Server.CreateObject("ADODB.Recordset")
		rsYKien.open sqlyk,con,1
		set rsYKien = nothing
	
		Response.Write	"<script language=""JavaScript"">" & vbNewline &_
		"	<!--" & vbNewline &_
		"		window.opener.location.reload();" & vbNewline &_
		"		window.close();" & vbNewline &_
		"	//-->" & vbNewline &_
		"</script>" & vbNewline		
end select
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <script type="text/javascript" src="../../ckeditor/ckeditor.js"></script>
    <script type="text/javascript" src="../../ckfinder/ckfinder.js"></script>
    <script type="text/javascript" src="../../js/jquery.js"></script>
    <link type="text/css" href="../../css/sweetalert.css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/sweetalert.min.js"></script>
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
</head>
<body>
    <form name="form1" id="form1" method="post" action="ans.asp?action=add">
        <table border="0" class="w3-table w3-table-all">
            <tr>
                <td>Họ và tên: </td>
                <td>
                    <input name="txtName" type="text" id="txtName" value="" class="w3-input w3-border w3-round" /></td>
            </tr>
            <tr>
                <td>Email: </td>
                <td>
                    <input name="txtEmail" type="text" id="txtEmail" value="" size="80" class="w3-input w3-border w3-round"/></td>
            </tr>
            <tr>
                <td>Tiêu đề: </td>
                <td>
                    <input name="txtTitle" type="text" id="txtTitle" value="" size="80" class="w3-input w3-border w3-round"/></td>
            </tr>
            <tr>
                <td>Câu hỏi : </td>
                <td>
                    <textarea name="txtAns" cols="62" rows="4" id="txtAns" class="w3-input w3-border w3-round"></textarea>
                </td>
            </tr>
            <tr>
                <td>Trả lời: </td>
                <td>
                    <textarea name="txtQuestion" cols="80" rows="10" id="txtQuestion"></textarea></td>
            </tr>
            <tr>
                <td>Cho phép hiển thị:</td>
                <td>
                    <input name="isShow" class="w3-check" type="checkbox" id="isShow" value="1"></td>
            </tr>
            <tr>
                <td colspan="2">
                    <div>
                        <input type="button" name="btn_ans" id="btn_ans" class="w3-btn w3-red w3-round" value="Cập nhật" />
                    </div>
                </td>
            </tr>
        </table>
    </form>
    <script type="text/javascript">
        CKEDITOR.replace('txtQuestion');
    </script>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#btn_ans").click(function () {
                if ($('#txtName').val() == '') {
                    $('#txtName').focus();
                    swal("BQT", "Xin vui lòng nhập họ tên.");
                }
                else if ($('#txtEmail').val() == '') {
                    $('#txtEmail').focus();
                    swal("BQT", "Xin vui lòng nhập email.");
                }
                else if (!isEmail($('#txtEmail').val())) {
                    $('#F_Email').focus();
                    swal("BQT", "Sai định dạng email.vd: abc@gmail.com");
                }
                else if ($('#txtTitle').val() == '') {
                    $('#txtTitle').focus();
                    swal("BQT", "Vui lòng nhập tiêu đề.");
                }
                else if ($('#txtAns').val() == '') {
                    $('#txtAns').focus();
                    swal("BQT", "Vui lòng nhập phần trả lời.");
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
</body>
</html>

