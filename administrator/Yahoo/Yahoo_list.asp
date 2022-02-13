<%session.CodePage=65001%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->

<%
f_permission = administrator(false,session("user"),"m_human")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
	<head>
		<title><%=PAGE_TITLE%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <script  type="text/javascript" src="/administrator/inc/common.js"></script>
        <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
        <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
        <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
        <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
        <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
        <script src="/administrator/skin/script/ajax-asp.js"></script>
        <link href="/administrator/css/skin1.css" rel="stylesheet" />
        <link type="text/css" href="../css/testadmin.css" rel="stylesheet" />
        <link href="/administrator/css/uploadanh.css" rel="stylesheet" />
	</head>
<body >
<%
	Title_This_Page="Quản lý ->Danh sách hỗ trợ"
	Call header()
%>
<%
    iStatus	=	Request.QueryString("iStatus")
    idNhanvien =Request.QueryString("id")	
    idNhanvien=CLng(idNhanvien)
        sql="SELECT * from  SupportYahoo"
        set rsN=server.CreateObject("ADODB.Recordset")
	    rsN.open sql,con,1
        iCount	=	rsN.recordcount + 1
        btnName = "Thêm"
    If  iStatus	=	"edit" then
        sql="SELECT * from  SupportYahoo WHERE ID=" & idNhanvien
        set rs=server.CreateObject("ADODB.Recordset")
	    rs.open sql,con,1
        iCount	=	rs.recordcount+1
        If not rs.eof then
            nickname = trim(rs("hoten"))
            Ghichu = trim(rs("ghichu"))
            phone = trim(rs("mobile"))
            email = trim(rs("email"))
            zalo = trim(rs("idzalo"))
            SmallPictureFileName = trim(rs("picturepath"))
            btnName = "Cập nhật"
        End if
    End if         
%>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10" style="background: #f1f1f1">
          <div class="main">
            <form action="Yahoo_update.asp"  name="YAHOOLIST" method="post" ENCTYPE="multipart/form-data">
              <div class="form-group">
                <label for="fullname" class="form-label">Tên nhân viên:</label>
                <div class="ver-covered-input">
                    <input name="nickname" id="nickname" class="form-control" value="<%=nickname%>"  />
                </div>
              </div>
              <div class="form-group">
                <label for="Ghichu" class="form-label">Chức vụ:</label>
                <div class="ver-covered-input">
                    <input name="Ghichu" id="Ghichu" class="form-control" value="<%=Ghichu%>" />
                </div>
              </div>
              <div class="form-group">
                <label for="fullname" class="form-label">Số điện thoại:</label>
                <div class="ver-covered-input">
                    <input name="phone" id="phone" class="form-control" value="<%=phone%>" />
                </div>
              </div>
              <div class="form-group">
                <label for="email" class="form-label">Email:</label>
                <div class="ver-covered-input">
                    <input name="email" id="email" class="form-control" value="<%=email%>" />
                </div>
              </div>
              <div class="form-group">
                <label for="zalo" class="form-label">Zalo:</label>
                <div class="ver-covered-input">
                    <input name="zalo" id="zalo" class="form-control" value="<%=zalo%>" />
                </div>
              </div>
              <div class="row">
    <div class="col-12 col-md-2 form-label">Tải ảnh lên</div>
    <div class="col-12 col-md-9">
           <div class="row">
                        <div></div>
                        <div class="col-md-3 col-sm-6 image-item">
                        <strong class="w3-text-red">Ảnh Bìa<%=i%>:</strong>
                        <%
                        if SmallPictureFileName<>"" then%>
			                <br/><img src="<%=NewsImagePath&SmallPictureFileName%>"width='200' height='100'><br/>
                            <input type="checkbox" name="RemoveImage" id="RemoveImage" value="1" class="w3-check" /> 
                            <label class="lb_remove" for="RemoveImage"><strong>Xóa tất cả</strong></label><br />
                        <% else %>
                            <br/><img src="" alt="" width='200' height='100' id="coverImage"><br/>
                            <input type="checkbox" name="RemoveImage" id="RemoveImage" value="1" class="w3-check" /> 
                            <label class="lb_remove" for="RemoveImage"><strong>Xóa tất cả</strong></label><br />
                        <% end if %>
                        <a class="btn input_upload">
                            <span><i class="fa fa-cloud-upload"></i> Chọn file</span>
                            <input name="SmallPictureFileName" type="file" id="SmallPictureFileName" onchange="pushImage()">
                        </a>
                        <br/>
                        </div><!--col-md-3-->
              </div>
      </div>
      </div>
      <div class="form-group-2">
                <label for="fullname" class="form-label">Phê duyệt:</label>
                <div class="ver-covered-input">
                <div class="checkbox-div"><input type="checkbox" checked="checked" name="attach_product" id="attach_product" value="1"
                    class="ckeckbox-children"><span class="checkbox-text">Tiếp tục gửi tin để thiết lập</span></div>
                    <div class="centered-item">
                        <input class="form-submit" type="submit" name="Button" value="<% =btnName %>" />
                        <input class="form-submit" type="reset" name="Button" value="Hủy"  />
                            <%
                                    if iStatus  = "edit" then 
                                        %>
                                        <input type="hidden" name="iStatus" value="edit">
                                        <input type="hidden" name="idnhanvien" value=<%=idNhanvien%> />
                                        <%  
                                    else
                                        %>
                                        <input type="hidden" name="iStatus" value="add">
                                         <input type="Hidden" name="old_Id" value="0">
		                                <input name="iCount"  type="hidden" value="<%=iCount%>">
                                        <%    
                                          
                                    end  if
                                        %>
                    </div>
            </div>
          </form>
<%
    rsN.close
    set rsN=nothing    
%>
          </div>
          <!-- Bảng ghi nhân viên -->
          <% call Callcool()%>
    </div>
</div>
<%Call Footer()%> 
<script>
  function pushImage() {
        const coverImage = document.getElementById("coverImage")
        coverImage.src = URL.createObjectURL(event.target.files[0])
    }
</script>
</body>
</html>
<% sub Callcool()%>
        <table class="table table-bordered w3-hoverable">
                <tr class="w3-blue">
                    <td>Họ tên</td>
                    <th>Chức vụ</th>
                    <td>Phone</td>
                    <td>Email</td>
                    <th>Zalo</th>
                    <th>Xư lý</th>
                </tr>
                <%
                sql="SELECT * from  SupportYahoo"
                set rsN=server.CreateObject("ADODB.Recordset")
	            rsN.open sql,con,1
                 response.write(rsN.recordcount)
                 j = 1
                    Do while not rsN.EOF 
                        tnickname = trim(rsN("hoten"))
                        tGhichu = trim(rsN("ghichu"))
                        tphone = trim(rsN("mobile"))
                        temail = trim(rsN("email"))
                        tzalo = trim(rsN("idzalo"))
                        tid = trim(rsN("id"))
                %>
                <tr>
                    <td><% =tnickname%></td>
                    <td><% =tGhichu%></td>
                    <td><% =tphone%></td>
                    <td><% =Email%></td>
                    <td><% = tzalo%></td>
                    <td>
                            <a href="Yahoo_list.asp?iStatus=edit&id=<%=tid%>">
                            <img src="/administrator/images/icon_edit_topic.gif" width="15" height="15" border="0" title="Sửa"></a>
                             <a href= "javascript: winpopup('/administrator/Yahoo/support_delete.asp','iStatus=delete&id=<%=tid%>',400,220);">
                            <img src="/administrator/images/icon_closed_topic.gif" width="15" height="15" border="0" title="Xóa"></a>
                    </td>
                </tr>
                <%
                    rsN.MoveNext
                    j= j+1
                    Loop
                %>
                <%	
                set rsN=nothing             
            %>
            </table>
            <% end sub%>
