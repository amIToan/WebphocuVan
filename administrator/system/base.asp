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
            key_ =  Request.QueryString("_key")
        %>
    </div>

    <div class="container-fluid">
        <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <%
                IF key_ = "base-add" THEN
            %>
            <form name="fcompany" id="fcompany" class="form-horizontal" method="post" enctype="multipart/form-data">
                <input name="_key" value="Add" type="hidden" />
                <div class="form-group text-center">
                    <h4>
                        <br />
                        THÊM CƠ SỞ </h4>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Tên cở sở:</label>
                    <div class="col-sm-10">
                        <input name="company" type="text" id="company" value="<%=csName_ %>" class="form-control" maxlength="100">
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Icon</label>
                    <div class="col-sm-10">
                        <input name="icon" type="file" id="icon">
                        bắt buộc file *.ico 
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Logo</label>
                    <div class="col-sm-10">
                        <input name="Logo" type="file" id="logo" size="25">
                        *.png, *.jpeg, *.gif, *.jpg
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Logo Footer</label>
                    <div class="col-sm-10">
                        <input name="logoF" type="file" id="logoF" size="25">
                        *.png, *.jpeg, *.gif, *.jpg 
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Địa chỉ</label>
                    <div class="col-sm-10">
                        <input name="address" type="text" class="form-control" maxlength="500" value="">
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Điện Thoại</label>
                    <div class="col-sm-10">
                        <input name="Tel" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Hotline:</label>
                    <div class="col-sm-10">
                        <input name="Hotline" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Thư điện tử (e-mail):</label>
                    <div class="col-sm-10">
                        <input name="Email" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Website:</label>
                    <div class="col-sm-10">
                        <input name="Website" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Mã số thuế:</label>
                    <div class="col-sm-10">
                        <input name="masothue" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Giấy phép KD:</label>
                    <div class="col-sm-10">
                        <input name="GPKD" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">T.G làm việc:</label>
                    <div class="col-sm-10">
                        <input name="calltime" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Mã google:</label>
                    <div class="col-sm-10">
                        <input name="idgoogle" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Fanpage Facebook:</label>
                    <div class="col-sm-10">
                        <input name="idfacebook" type="text" class="form-control" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Link Youtube:</label>
                    <div class="col-sm-10">
                        <input name="idyoutube" type="text" class="form-control" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Nick sky:</label>
                    <div class="col-sm-10">
                        <input name="idskype" type="text" class="form-control" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Twiter:</label>
                    <div class="col-sm-10">
                        <input name="idTwiter" type="text" class="form-control" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Link G+:</label>
                    <div class="col-sm-10">
                        <input name="idgplus" type="text" class="form-control" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Meta Title:</label>
                    <div class="col-sm-10">
                        <input name="page_title" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Meta description:</label>
                    <div class="col-sm-10">
                        <input name="meta_description" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Meta keyword:</label>
                    <div class="col-sm-10">
                        <input name="meta_keywords" type="text" class="form-control" maxlength="100" value="" />
                    </div>
                </div>


                <div class="form-group">
                        <div class="col-md-4 col-md-offset-2  text-left">
                            <input id="" type="checkbox" name="ckeckmain" value="1"  /> Mặc định.
                        </div>
                         <div class="col-md-4 text-right">
                            <input id="btnsubmit" type="button" name="Submit" value="Tạo mới" class="btn btn-primary" />
                        </div>
                    </div>
            </form>
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
            <%  ELSEIF key_ = "Edit" THEN %>
            <%
                id_ = Trim(Replace(Request.QueryString("_id"),".html","")) 
                IF NOT IsEmpty(id_) AND IsNumeric(id_) THEN
                    sqlEdit = "SELECT * FROM Company WHERE  ID = '"&id_&"'"
                    set rsEdit = Server.CreateObject("ADODB.Recordset")
                    rsEdit.open sqlEdit,con,1
                    IF NOT rsEdit.EOF THEN
                        csID_      = rsEdit("ID")
                        csName_    = rsEdit("company")
                        csicon_    = rsEdit("icon")
                        csLogo_    = rsEdit("Logo")
                        csLogoF_    = rsEdit("LogoF")
                        IF csicon_ <> "" OR csicon_ <> "NULL" THEN
                            iconM = "<img src='/images/logo/"&csicon_&"'  style='max-width:100px;' />"
                        ELSE
                            iconM =""
                        END IF
                        IF csLogo_ <> "" OR csLogo_ <> "NULL" THEN
                            logoM = "<img src='/images/logo/"&csLogo_&"'  style='max-width:100px;' />"
                        ELSE
                            logoM =""
                        END IF
                        IF csLogoF_ <> "" OR csLogoF_ <> "NULL" THEN
                            LogoF = "<img src='/images/logo/"&csLogoF_&"'  style='max-width:100px;' />"
                        ELSE
                            LogoF =""
                        END IF                       
                        csTel_           = rsEdit("Tel")
                        csHotline_       = rsEdit("Hotline")
                        csAddress_       = rsEdit("address")
                        csEmail_         = rsEdit("Email")
                        csWebsite_       = rsEdit("Website")
                        csMasothue_      = rsEdit("Masothue")
                        csGPKD_          = rsEdit("GPKD")
                        cspage_title_    = rsEdit("page_title")             
                        csmeta_desc_     = rsEdit("meta_description")
                        csmeta_keywords_ = rsEdit("meta_keywords")
                        csidgoogle_      = rsEdit("idgoogle")
                        csDesc_          = rsEdit("introduction")
                        csshow_intro_    = rsEdit("show_intro_home")                    
                        csidgplus_       = rsEdit("idgplus")
                        csidyoutube_     = rsEdit("idyoutube")
                        csidskype_       = rsEdit("idskype")
                        csidfacebook_    = rsEdit("idfacebook")
                        cslang_          = rsEdit("lang")
                        csshow_          = rsEdit("show")
                        IF csshow_ <> "" AND IsNumeric(csshow_) THEN
                            idShow_ = " checked "
                        ELSE
                            idShow_ = ""
                        END IF
                
                        cscalltime_      = rsEdit("calltime")
                        csidtwiter_      = rsEdit("idtwiter")
                                      
                
            %>
            <div class="col-md-10">
                <form name="F_baseEdit" id="F_baseEdit" class="form-horizontal" method="post" enctype="multipart/form-data">
                    <input name="_key" value="Update" type="hidden" />
                    <input name="_id" value="<%=id_ %>" type="hidden" />
                    <input name="_Ficon" value="<%=csicon_ %>" type="hidden" />
                    <input name="_FLogo" value="<%=csLogo_ %>" type="hidden" />
                    <input name="_FlogoF" value="<%=csLogoF_ %>" type="hidden" />
                    <div class="form-group text-center">
                        <h4>
                            <br />
                            SỬA THÔNG TIN </h4>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Tên cở sở:</label>
                        <div class="col-sm-10">
                            <input name="FEdit_copany" type="text" id="FEdit_copany" value="<%=csName_ %>" class="form-control" maxlength="100">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Icon</label>
                        <div class="col-sm-10">
                            <%=iconM %>
                            <input name="icon" type="file" id="FEdit_icon">
                            bắt buộc file *.ico
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Logo</label>
                        <div class="col-sm-10">
                            <%=logoM %>
                            <input name="Logo" type="file" id="FEdit_logo" size="25">
                            *.png, *.jpeg, *.gif, *.jpg 
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Logo Footer</label>
                        <div class="col-sm-10">
                            <%=LogoF %>
                            <input name="logoF" type="file" id="FEdit_logoF" size="25">
                            *.png, *.jpeg, *.gif, *.jpg 
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Địa chỉ</label>
                        <div class="col-sm-10">
                            <input name="address" type="text" class="form-control" maxlength="500" value="<%=csAddress_ %>">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Điện Thoại</label>
                        <div class="col-sm-10">
                            <input name="Tel" type="text" class="form-control" maxlength="100" value="<%=csTel_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Hotline:</label>
                        <div class="col-sm-10">
                            <input name="Hotline" type="text" class="form-control" maxlength="100" value="<%=csHotline_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Thư điện tử (e-mail):</label>
                        <div class="col-sm-10">
                            <input name="Email" type="text" class="form-control" maxlength="100" value="<%=csEmail_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Website:</label>
                        <div class="col-sm-10">
                            <input name="Website" type="text" class="form-control" maxlength="100" value="<%=csWebsite_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Mã số thuế:</label>
                        <div class="col-sm-10">
                            <input name="masothue" type="text" class="form-control" maxlength="100" value="<%=csMasothue_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Giấy phép KD:</label>
                        <div class="col-sm-10">
                            <input name="GPKD" type="text" class="form-control" maxlength="100" value="<%=csGPKD_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">T.G làm việc:</label>
                        <div class="col-sm-10">
                            <input name="calltime" type="text" class="form-control" maxlength="100" value="<%=cscalltime_ %>" />
                        </div>
                    </div>
                     <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Mô Tả:</label>
                        <div class="col-sm-10">
                            <input name="csDesc" type="text" class="form-control" maxlength="100" value="<%=csDesc_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Mã google:</label>
                        <div class="col-sm-10">
                            <input name="idgoogle" type="text" class="form-control" maxlength="100" value="<%=csidgoogle_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Fanpage Facebook:</label>
                        <div class="col-sm-10">
                            <input name="idfacebook" type="text" class="form-control" value="<%=csidfacebook_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Link Youtube:</label>
                        <div class="col-sm-10">
                            <input name="idyoutube" type="text" class="form-control" value="<%=csidyoutube_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Nick sky:</label>
                        <div class="col-sm-10">
                            <input name="idskype" type="text" class="form-control" value="<%=csidskype_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                    <label for="inputEmail3" class="col-sm-2 control-label">Twiter:</label>
                    <div class="col-sm-10">
                        <input name="idTwiter" type="text" class="form-control" value="<%=csidtwiter_ %>" />
                    </div>
                </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Link G+:</label>
                        <div class="col-sm-10">
                            <input name="idgplus" type="text" class="form-control" value="<%=csidgplus_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Meta titel:</label>
                        <div class="col-sm-10">
                            <input name="page_title" type="text" class="form-control" maxlength="100" value="<%=cspage_title_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Meta description:</label>
                        <div class="col-sm-10">
                            <input name="meta_description" type="text" class="form-control" maxlength="1000" value="<%=csmeta_desc_ %>" />
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="inputEmail3" class="col-sm-2 control-label">Meta keyword:</label>
                        <div class="col-sm-10">
                            <input name="meta_keywords" type="text" class="form-control" maxlength="1000" value="<%=csmeta_keywords_ %>" />
                        </div>
                    </div>

                    <div class="form-group " >
                        <div class="col-md-4 col-md-offset-2  text-left">
                            <input id="" type="checkbox" name="ckeckmain" value="1"  <%=idShow_ %>   /> Mặc định.
                        </div>
                         <div class="col-md-4 text-right">
                           <input id="Btn_Editsubmit" type="button" name="Submit" value="Cập nhật" class="btn btn-primary" />
                        </div>
                    </div>






                </form>
            </div>
            <script type="text/javascript">
                $("#Btn_Editsubmit").click(function () {
                    if ($('#FEdit_copany').val() == '') {
                        $('#FEdit_copany').focus();
                        swal("BQT", "Hãy nhập tên cơ sở.");
                    }
                    else {
                        Navi_baseUpdate('', '0');
                    }
                });
                function isEmail(email) {
                    var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                    return regex.test(email);
                }
            </script>
            <%  
                    END IF ' Not data
                    ELSE 
            %>
            Lỗi.
            <%  END IF 'EMPTY EDIT %>
            
            <%  END IF %>
        </div>
    </div>


    <%Call Footer()%>
<script src="/administrator/skin/script/sweetalert.min.js"></script>
</body>

</html>
