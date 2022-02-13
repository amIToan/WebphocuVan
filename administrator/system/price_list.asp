<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
if request.QueryString("lang")="EN" then
	lang="EN"
else
	lang="VN"
end if
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>	
    <link href="../../css/styles.css" rel="stylesheet" type="text/css"></head>
<link href="../../css/CommonSite.css" rel="stylesheet" />
    <script src="/ckeditor/ckeditor.js"></script>
<script src="/ckfinder/ckfinder.js"></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Title_This_Page="Quản lý -> cập nhận nhà xuất bản"
	Call header()
	Call Menu()
	
%>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="right" valign="top"> 
    <td height="25"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
		<a href="javascript: winpopup('cat_chooselang.asp','<%=lang%>',220,120);"> 
      Chọn Ngôn ngữ</a> 
    </td>
  </tr>
</table>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM XSEOPrice ORDER BY ID"
    i=1
	rs.open sql,con,1
%>
<form action="price_update.asp?action=update"  name="NXBLIST" method="post">
<table width="800px"  align="center" cellpadding="0" cellspacing="0"  class="CTxtContent">
  <%
  i=1
  Do while not rs.eof
  %>
	<tr>
    <td>
        <div class =" border-box-text" style="padding:5px; <%if i mod 2> 0 then%>background-color:#ffd800;<%end if%>">
            <span class="CTieuDe" style="color:#2957A4;">Biểu giá <%=i%></span><br />
        Tiêu đề<br />
	<input name="id<%=i%>" type="hidden" value="<%=rs("id")%>">

    <input name="Title<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Title")%>" size="100"><br />
    Hiển thị ghi chú:<br />

    <textarea name="Note<%=i%>" id="Note<%=i%>" cols="50" rows="5"><%=rs("Note")%></textarea><br />
        Đường dẫn cho khách hàng download: <br />
	<input name="Download<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Download")%>" size="100"><br />
        Giá: <br />
	<input name="Price<%=i%>" type="text" class="CTextBoxUnder" onKeyUp="javascript: DisMoneyThis(this);" value="<%=Dis_str_money(rs("Price"))%>" size="30" style="text-align:right;"> VNĐ<br />
            Giá khi có mã giản:<br />
	<input name="PriceOff<%=i%>" type="text" class="CTextBoxUnder" onKeyUp="javascript: DisMoneyThis(this);" value="<%=Dis_str_money(rs("PriceOff"))%>" size="30" style="text-align:right;"> VNĐ<br />
        Số tháng được kích hoạt:<br />
    <input name="Month<%=i%>" type="text" class="CTextBoxUnder" value="<%=rs("Month")%>" size="30" onKeyUp="javascript: DisMoneyThis(this);"> <br />
        Mô tả:<br />
             <textarea name="Description<%=i%>" id="Description<%=i%>" cols="50" rows="30"><%=rs("Description")%></textarea><br />
            <div class="ButtonCircle" style="width:64px; height:64px;text-align:center; background-color:#ff6a00; " >
	<%
		Response.Write	"<a href=""javascript: winpopup('price_update.asp','" & lang & "&id=" & rs("id") & "&action=del',300,150);"" style=""text-decoration:none;vertical-align:middle;"" class=""CTieuDeNho"" color=""#FFF"">Xóa</a>"
	%>	
                </div>
        <script type="text/javascript">
            CKEDITOR.replace('Description<%=i%>');
            var editor = CKEDITOR.replace('Description<%=i%>');
            CKFinder.setupCKEditor(editor, '/ckfinder/');
        </script>
    </div>
    </td>
  </tr>	
	<tr>
    <td style="height:50px;">
    </td>
	</tr>    	
  <%
  i=i+1
  rs.movenext
  Loop
  rs.close
  set rs=nothing
  %>
  
  <tr >
    <td>
        <div class =" border-box-text" style="padding:5px; <%if i mod 2> 0 then%>background-color:#ffd800;<%end if%>">
            <span class="CTieuDe" style="color:#2957A4;">Biểu giá <%=i%> - MỚI (Thêm mới nếu có nội dung)</span><br />
        Tiêu đề<br />
	<input name="id<%=i%>" type="hidden" value="">
    <input name="Title<%=i%>" type="text" class="CTextBoxUnder" value="" size="100"><br />
    Hiển thị ghi chú:<br />
    <textarea name="Note<%=i%>" id="Note<%=i%>" cols="50" rows="5"></textarea><br />
        Đường dẫn cho khách hàng download: <br />
	<input name="Download<%=i%>" type="text" class="CTextBoxUnder" value="" size="100"><br />
        Giá: <br />
         <input name="Price<%=i%>" type="text" class="CTextBoxUnder" value="" size="30" onKeyUp="javascript: DisMoneyThis(this);" style="text-align:right;"> VNĐ<br />
                       Giá khi có mã giản:<br />
	    <input name="PriceOff<%=i%>" type="text" class="CTextBoxUnder" onKeyUp="javascript: DisMoneyThis(this);"  size="30" style="text-align:right;"> VNĐ<br />
 
        Số tháng được kích hoạt:<br />
    <input name="Month<%=i%>" type="text" class="CTextBoxUnder" value="" size="30" onKeyUp="javascript: DisMoneyThis(this);"> <br />
        Mô tả:<br />
             <textarea name="Description<%=i%>" id="Description<%=i%>" cols="50" rows="30"></textarea>
        <script type="text/javascript">
            CKEDITOR.replace('Description<%=i%>');
            var editor = CKEDITOR.replace('Description<%=i%>');
            CKFinder.setupCKEditor(editor, '/ckfinder/');
        </script>
	</td>
  </tr>	
</table>
<center>
	<input type="hidden" name="iCount" value="<%=i%>">
	<input type="hidden" name="action" value="update">
    <div class="CTieuDeNho" >
	<input name="submit" type="submit" id="submit" value="  Cập nhật  " class="ButtonCircle" style="width:72px; height:72px;">
        </div>
</center>
</form>
<%Call Footer()%>
</body>
</html>
