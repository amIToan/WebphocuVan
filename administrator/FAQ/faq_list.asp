<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_faq")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
f_permission = administrator(false,session("user"),"m_faq")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%IF Request.form("action")="Search" then
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	
	OrderType=ReplaceHTMLToText(Request.form("OrderType"))
ELSE
	Ngay1=Day(now())
	Thang1=Month(now())-1
	Nam1=Year(now())
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
END IF
%>
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
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
    <script type="text/javascript">
        function confirm_Del() {
            if (window.confirm("Bạn chắc chắn xóa không?") == true) {
                return true;
            } else {
                return false;
            }
        }
    </script>
</head>

<body>
    <div class="container-fluid">
        <%Call header()%>
    </div>
    <div class="container-fluid">
        <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
          <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fFAQ" onsubmit="return checkme();">
            <table align="center" cellpadding="0" cellspacing="0" width="770" class="w3-table w3-table-all w3-round w3-margin">
                                <tr>
                                    <td align="right" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
                                        <%
			                                Call List_Date_WithName(Ngay1,"DD","Ngay1")
			                                Call List_Month_WithName(Thang1,"MM","Thang1")
			                                Call List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
                                        %>
                                        <img src="../images/right.jpg" width="9" height="9" align="absmiddle">
                                        <%
			                                Call List_Date_WithName(Ngay2,"DD","Ngay2")
			                                Call List_Month_WithName(Thang2,"MM","Thang2")
			                                Call  List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
                                        %>
                                        <input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
                                        <input type="hidden" name="action" value="Search">
                                        <input type="hidden" name="OrderType" value="">
                                    </td>
                                </tr>
                            </table>
          </form>
<script  type="text/javascript">
    function order(OrderType) {
        if (!checkme())
            return;
        document.fThongke.OrderType.value = OrderType;
        document.fThongke.submit();
    }
    function checkme() {
        if (document.fThongke.Ngay1.value == 0) {
            alert("Bạn chưa chọn ngày!");
            document.fThongke.Ngay1.focus();
            return false;
        }
        if (document.fThongke.Thang1.value == 0) {
            alert("Bạn chưa chọn tháng!");
            document.fThongke.Thang1.focus();
            return false;
        }
        if (document.fThongke.Nam1.value == 0) {
            alert("Bạn chưa chọn năm!");
            document.fThongke.Nam1.focus();
            return false;
        }
        if (document.fThongke.Ngay2.value == 0) {
            alert("Bạn chưa chọn ngày!");
            document.fThongke.Ngay2.focus();
            return false;
        }
        if (document.fThongke.Thang2.value == 0) {
            alert("Bạn chưa chọn tháng!");
            document.fThongke.Thang2.focus();
            return false;
        }
        if (document.fThongke.Nam2.value == 0) {
            alert("Bạn chưa chọn năm!");
            document.fThongke.Nam2.focus();
            return false;
        }
        return true;
    }
</script>
            <%Call Faq_list()  %>
        </div><!---/.col-md-10--->
    </div>

<%sub  Faq_list()
FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
IF Request.form("action")="Search" and IsDate(FromDate) and IsDate(ToDate) THEN
	Dim rsYKien
	Set rsYKien=Server.CreateObject("ADODB.Recordset")
	sqlYKien="SELECT Top 500 * from Y_KIEN "&_
		"where (DATEDIFF(dd, ngaytao, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, ngaytao, '" & ToDate & "') >= 0 ) "
	sqlYKien=sqlYKien+	" order by ID Desc"		
    'response.write sqlYKien
	rsYKien.open sqlYKien,con,3
	i=0

    %>
    <table border="0" class="CTxtContent w3-table w3-table-all" style="border-spacing: 0px; width: 100%;" >
        <tr>
            <td>
                <a href="javascript: winpopup('ans.asp','0&action=new',800,600);" class="CSubMenu w3-btn w3-blue w3-round w3-right"><i class="fa fa-plus-square" aria-hidden="true"></i> Thêm mới</a>
            </td>
        </tr>
    </table>
    <%

	if not rsYKien.eof then
	iSTT=0
    %>
    <table border="0" class="CTxtContent w3-table w3-table-all">
        <tr class="w3-blue">
            <th><strong>Tên khách hàng</strong></th>
            <th><strong>Tiêu đề</strong></th>
            <th><strong>Số Đ.Thoại</strong></th>
            <th>Email</th>
            <th><strong>Ngày tạo</strong></th>
            <th><strong>Hiển thị</strong></th>
            <th>Chức năng</th>
        </t>
        <%
  	do while not rsYKien.eof 
		iSTT=iSTT+1
		id=Trim(rsYKien("id"))
		tel=rsYKien("tel")
		hovaten=Trim(rsYKien("hovaten"))
		noidung=Trim(rsYKien("noidung"))
		faq=rsYKien("faq")
		show=rsYKien("show")
		ngaytao=rsYKien("ngaytao")
		Traloi	=	rsYKien("Traloi")
		Title	=	rsYKien("tieude")
        Title=Mid(Title,1,40)
        %>
        <tr>
            <td class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>">
                <%=hovaten%><br>
                <span style="color: #ba5a16;"><%=tieude%>		</span></td>
            <td style="<%=setStyleBorder(0,1,0,1)%>"><span style="color: #bd6418;"><%=Title%></span>
            <td style="<%=setStyleBorder(0,1,0,1)%>"><span style="color: #bd6418;"><%=tel%></span></td>
            <td> 
                <a class="CSubMenu" title="Email: <%=rsYKien("Email")%>" href="javascript: winpopup('../FAQ/sendmail.asp','<%=id%>&action=send',990,600);">
                    <%=rsYKien("email")%>
                </a>
            </td>
            <td style="<%=setStyleBorder(0,1,0,1)%>; text-align: right;">
                <%=Day(ConvertTime(ngaytao))%>/
			    <%=Month(ConvertTime(ngaytao))%>/
			    <%=Year(ConvertTime(ngaytao))%>	</td>
            <td style="<%=setStyleBorder(0,1,0,1)%>; text-align: center;">
                <%
			if show = 0  then
                %>
                <img src="../images/icon-deactivate.gif" width="16" height="16" border="0" alt="Hiển thị">
                <%			
                    else
                %>
                <img src="../images/icon-activate.gif" width="16" height="16" border="0" alt="Không hiển thị">
                <%			
                    end if
                %> 
            </td>
            <td style="<%=setStyleBorder(0,1,0,1)%>; text-align: center;">

                 <a class="w3-btn w3-blue w3-round" href="javascript: winpopup('ykien.asp','0&id=<%=id%>&action=edit',800,600);">
                    <font size="2" face="Arial, Helvetica, sans-serif"><i class="fa fa-pencil-square-o" aria-hidden="true"></i> Xem</font>
                </a>                
                <a class="w3-btn w3-red w3-round" href="ykien.asp?id=<%=id%>&action=del" target="_blank" onclick="return confirm_Del();">            
                    <font size="2" face="Arial, Helvetica, sans-serif"><i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Xóa</font>
                </a>
            </td>

        </tr>
        <%
  	i=i+1
	rsYKien.movenext
    Loop
        %>
    </table>
    <%
    rsYKien.close
    set rsYKien=nothing
	else
		Response.Write("<center>Không có nội dung để hiển thị </center>")
	end if

End if


end sub
    %>
</body>
</html>











