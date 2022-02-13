<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
    lang = Session("Language")
    if lang = "" then lang = "VN"
        Session.Timeout = 1440
        f_permission = administrator(false,session("user"),"m_editor")
    if f_permission <= 0 then
	    response.Redirect("/administrator/info.asp")
    end if

iStatus	=	Request.QueryString("iStatus")

redim PictureFile(16)
redim ContentPicture(16)
if iStatus	=	"add" then
	CatId=0
	strTitle	=	"<img src=""../../images/icons/icon_key_points.gif"" width=48 height=48 align=""absmiddle"">  Nhập "& GetNameOfCategoryLoai(CategoryLoai) &" mới"
	Title_This_Page="Quản lý -> Nhập "&GetNameOfCategoryLoai(CategoryLoai)&" mới."
	PictureId=0
	Price=0
	PriceNet=0
	price_usd	=	0		
	StoreOf=""
	Unit = "đ"
	maker=" "
	ck	=	0	
    Store = "" 
    Size=0
	Weight =0

else
	strTitle	=	"<img src=""../../images/icons/icon_but.jpg"" width=45 height=35 align=""absmiddle""> Sửa "&GetNameOfCategoryLoai(CategoryLoai)
	Title_This_Page="Quản lý -> Sửa "&GetNameOfCategoryLoai(CategoryLoai)
	NewsId=Request.QueryString("NewsId")
	CatId=Request.QueryString("CatId")
	if not IsNumeric(NewsId) or not IsNumeric(CatId) then
		response.Redirect("/administrator/")
		response.End()
	else
		NewsId=CLng(NewsId)
		CatId=CLng(CatId)
	end if
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ed")
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	'Lấy ra các CatId có liên hệ với tin đồng thời thuộc phạm vi quyền của User.
	sql1=GetSQL_For_Search(session("LstCat"),session("LstRole"),session("user"),"NONE")
	sql="SELECT d.CategoryId from News n, NewsDistribution d"
	sql=sql & " WHERE d.NewsId=" & Newsid & " and d.NewsId=n.NewsId "
   ' Response.Write sql
	categoryid=""
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		CatNum=0
		Do while not rs.eof
			CatNum=CatNum+1
			categoryid=categoryid & " " & rs("CategoryId")
		rs.movenext
		Loop
	rs.close
	
	if CatNum=1 then
	'Nếu số lượng chuyên mục liên quan với tin là một
		scategoryid=""
    isMulti = 0
	else
    isMulti =1
		scategoryid=trim(categoryid)
	end  if

	'Lấy thông tin từ record có NewsId=NewsId
	sql="SELECT * from News WHERE LanguageID = '"&lang&"' and NewsId=" & NewsId

	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	iF not rs.eof then
		
        config=Trim(rs("Author"))

		IDCode=rs("IDCode")

		PublicationNo=CLng(rs("PublicationNo"))


		if rs("IsHomeNews") then
			IsHomeNews=1
		else
			IsHomeNews=0
		end if

		if rs("IsCatHomeNews") then
			IsCatHomeNews=1
		else
			IsCatHomeNews=0
		end if
		
		if rs("AdsHome") then
			AdsHome=1
		else
			AdsHome=0
		end if	
		if rs("AdsNews") then
			AdsNews=1
		else
			AdsNews=0
		end if	
    	if rs("IsHotNews") then
			IsHotNews=1
		else
			IsHotNews=0
		end if	
        trongluong=	rs("EmptyStore")
		Title=Trim(rs("Title"))
        url_video = Trim(rs("url_video"))
        IDCode = Trim(rs("IDcode"))
        Source = Trim(rs("Source"))
		desc=Trim(rs("Description"))
        Tskt =Trim(rs("DecsBannerImage"))
		mota=Replace(mota,"<br>",chr(13) & chr(10))   
        bodyx=Trim(rs("body"))
		PictureId =Clng(rs("PictureId"))
		PictureAlign=Trim(rs("PictureAlign"))
		Author=trim(rs("Author"))
		Source=trim(rs("Source"))
		StatusId=Trim(rs("StatusId"))
		Price=Clng(rs("Price"))
		PriceNet=Clng(rs("PriceNet"))
		Unit	=	trim(rs("Unit"))
		if PriceNet<>0 then
		discount	=	100*(PriceNet-Price)/PriceNet		
		else
		discount = 0
		end if
		StoreOf=trim(rs("StoreOf"))
			
		Size=rs("Size")
		Weight =rs("Weight")
		EmptyStore=	rs("EmptyStore")
        meta_keyword =	rs("meta_keyword")
		meta_desc=	rs("meta_desc")

    datet = rs("CreationDate")
      DateNow = Day(datet)&"-"&Month(datet)&"-"&Year(datet)&" "&Hour(datet)&":"&Minute(datet)
		set rs=nothing
		if PictureId<>0 then
			sql="select * FROM Picture WHERE PictureId=" & PictureId
			set rs=server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1				
                if not rs.EOF then
				    SmallPictureFileName=rs("SmallPictureFileName")
				    LargePictureFileName=rs("LargePictureFileName")
				    PictureAuthor=rs("PictureAuthor")

                    For n=1 to 16
				        PictureFile(n)      =	rs("PictureFile"&n)
				        ContentPicture(n)   =   rs("ContentPicture"&n)
				    Next
                end if
			rs.close
		else
			PictureCaption=""
			LargePictureFileName=""
			PictureAuthor=""
		end if
		set rs=nothing
	end if	
end if

    if iStatus  = "edit" then       
       btnName = "Cập nhật"         
       if LargePictureFileName <> "" then
           LoadAdsImg = "<img src='/images_upload/"&LargePictureFileName&"'  style='width:600px;'/> <br />"
       else
            LoadAvata = ""
       end if  
       LoadAvata = "<img src='/images_upload/"&SmallPictureFileName&"'  style='width:200px;'/> <br />" 
    else
       btnName = "Đăng tin"
       LoadAvata = ""
       LoadAdsImg = ""
    end  if
  
    if Trim(DateNow) = "" then DateNow = Day(now())&"-"&Month(now())&"-"&Year(now())&" "&Hour(now())&":"&Minute(now())
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript"  src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">    
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <script src="/ckeditor/ckeditor.js" type="text/javascript"></script>
    <script src="/ckfinder/ckfinder.js" type="text/javascript"></script>
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script type="text/javascript" src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script type="text/javascript" src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../css/testadmin.css" rel="stylesheet" />
    <link href="/administrator/css/uploadanh.css" rel="stylesheet" />
</head>
<body>
    <div class="container-fluid">
        <%
            Call header() 
            key_ =  Request.QueryString("_key")
        %>
    </div>
    <div class="container-fluid ">
        <div class="row" style="background: #f1f1f1">
        <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10" style="background: #f1f1f1">
            <div class="main">
                 <form name="fInsert" enctype="multipart/form-data" method="post">
            <div class="form-group-2">
                <label for="fullname" class="form-label">Chuyên mục:</label>
                <div class="ver-covered-input">
                    <%Call List_Category_Depend_Role(CatId, "Lựa chọn","NONE",Session("LstRole"),"ed",0,isMulti)%>
                    <input style="display: none;" type="text" name="categoryid" id="categoryid" value="<%=scategoryid %>" />
                </div>
            </div>
            <div class="form-group-2">
                <label for="fullname" class="form-label">Điều hướng:</label>
                <div class="hor-covered-input">
                    <div class="checkbox-div">
                        <input type="checkbox" name="IsHomeNews" id="IsHomeNews" value="1" <%if IsHomeNews <> 0 then Response.Write("checked")%>
                            class=" ckeckbox-children"><span class="checkbox-text">Trang chủ</span></div>
                    <div class="checkbox-div">
                        <input type="checkbox" name="AdsNews" id="" value="1" <%if AdsNews <> 0 then Response.Write("checked")%>
                            class=" ckeckbox-children"><span class="checkbox-text">Liên quan</span>
                        </div>
                </div>
            </div>
            <div class="form-group-2">
                <label for="fullname" class="form-label">Mô tả:</label>
                <div class="ver-covered-input">
                    <span class="checkbox-text">Tiêu đề chính (*)</span>
                    <input name="Title" class="form-control" type="text" id="Title" maxlength="200" value="<%=Title%>" >
                    <span class="checkbox-text">Mô tả ngắn:</span>
                    <textarea name="F_Tskt""  id="F_Tskt" class="form-control" placeholder="Write something.." style="height:100px"><%=Tskt%></textarea>
                    <div class="sub-hor-covered-input">
                        <div>
                            <label for="f_Price" class="checkbox-text">Giá bán:</label>
                            <input id="f_Price" name="f_Price" type="text" placeholder="" class="form-control" onkeyup="javascript: DisMoneyThis(this);" maxlength="50">
                        </div>
                        <div>
                            <label for="f_unit" class="checkbox-text">Đơn vị:</label>
                            <!--<input id="giaban" name="giaban" type="text" placeholder="" class="form-control">-->
                            <input name="f_unit" type="text" id="f_unit" class="form-control" value="<%=Unit %>" class="form-control" />
                        </div>
                        <div>
                            <label for="f_Size" class="checkbox-text">1 đơn vị diện tích:</label>
                           <!-- <input id="giaban" name="giaban" type="text" placeholder="" class="form-control">-->
                            <input id="f_Size" class="form-control" maxlength="50" name="f_Size" type="text" value="<%=Size%>" />
                        </div>
                        <div>
                            <label for="f_Speed" class="checkbox-text">Tiến độ:</label>
                          <!--  <input id="giaban" name="giaban" type="text" placeholder="" class="form-control">-->
                            <input id="f_Speed" type="text" class="form-control" name="f_Weight" value="<%=Weight %>">
                        </div>
                    </div>
                </div>
            </div>
            <div class="form-group-2">
                <label for="fullname" class="form-label">Link video:</label>
                <div class="ver-covered-input">
                    <span class="checkbox-text">Truy cập video trên youtube.com (*)</span>
                    <input name="url_video" id="url_video" type="text" placeholder="" class="form-control" value="<%=url_video%>">
                    <!--<!--<span class="checkbox-text">Mô tả:</span>
                    <textarea name="Fconfig"  id="Fconfig" class="form-control" placeholder="Write something.." style="height:100px"> <%=config%></textarea>-->
                    <span class="checkbox-text">Nội dung tin:</span>
                    <textarea id="bodyx" name="bodyx" placeholder="Write something.." style="height:200px"
                        class="form-control"> <%=bodyx%> </textarea>
                </div>
            </div>
                     <!-- Phan up anh -->
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
                            <br/><img src="/administrator/images/images.png" alt="" width='200' height='100' id="coverImage"><br/>
                            <input type="checkbox" name="RemoveImage" id="RemoveImage" value="1" class="w3-check" /> 
                            <label class="lb_remove" for="RemoveImage"><strong>Xóa tất cả</strong></label><br />
                        <% end if %>
                        <a class="btn input_upload">
                            <span><i class="fa fa-cloud-upload"></i> Chọn file</span>
                            <input name="SmallPictureFileName" type="file" id="SmallPictureFileName" onchange="pushImage()">
                        </a>
                        <br />
                        </div><!--col-md-3-->
                    <%
	                For i =1 to 7
		                PictureFile1	=	PictureFile(i)
                        if PictureFile1<>"" then
                            Filetype = Right(PictureFile1,len(PictureFile1)-Instr(PictureFile1,"."))
                        end if

                        if i=1 then
                            css_=""
                        else
                            if (PictureFile1<>"") then
                                css_=""
                            else
                                css_="style='display:none;'"
                            end if
                        end if
                    %>
                    <div class="col-md-3 col-sm-6 image-item" id="Picture<%=i%>T" <%=css_ %> >
                    <strong class="w3-text-red">Ảnh <%=i%>:</strong>
                    <div class="area-upload">
                        <%if PictureFile1<>"" And (Lcase(Filetype)="jpg" Or Lcase(Filetype)="gif" Or Lcase(Filetype)="jpeg" Or Lcase(Filetype)="png") then%>
			                <img id="imgFile<%=i%>" src="<%=NewsImagePath&PictureFile1%>"width='200' height='100'><br/>
                            <input name="PictureDel<%=i%>" id="PictureDel<%=i%>" type="checkbox" value="1" class="w3-check"> 
                            <label for="PictureDel<%=i%>"><strong>Xóa ảnh</strong></label><br />
                        <%else %>
                            <img id="imgFile<%=i%>" src="/administrator/images/images.png" width='200' height='100'><br/>
                            <input name="PictureDel<%=i%>" id="PictureDel<%=i%>" type="checkbox" value="1" class="w3-check"> 
                            <label class="lb_remove" for="PictureDel<%=i%>"><strong>Xóa ảnh</strong></label><br />
                        <% end if %>
                        <a class="btn input_upload">
                            <span><i class="fa fa-cloud-upload"></i> Chọn file</span>
                            <input class="custom-file " name="PictureFile<%=i%>" type="file" id="PictureFile<%=i%>" onchange="javascript: OnBrowser(this,<%=i%>)">                
                        </a><br />
                    </div>
                    </div>
                    <%Next%>
                                        </div><!--row-->
    </div>
</div>
<link href="//cdn.quilljs.com/1.3.6/quill.snow.css" rel="stylesheet">
    <script type="text/javascript" src="//cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/highlight.min.js"></script>
    <script type="text/javascript" src="//cdn.quilljs.com/1.3.6/quill.min.js"></script>

<script type="text/javascript">
    function OnBrowser(OjThis, iNum) {
        // load src into images
        console.log(OjThis, iNum)
        var image = document.getElementById('imgFile' + iNum);
        image.src = URL.createObjectURL(event.target.files[0]);

        str = OjThis.value
        iNum = iNum + 1
        console.log(str);
        if (str != '') {
            document.getElementById("Picture" + iNum + "T").style.display = "";
        } else {
            for (i = iNum; i <= 16; i++) {
                document.getElementById("Picture" + i + "T").style.display = "none";
                strtemp = "document.fInsert.PictureFile" + i + ".value='';";
                eval(strtemp);
            }
        }
    }
    function pushImage() {
        const coverImage = document.getElementById("coverImage")
        coverImage.src = URL.createObjectURL(event.target.files[0])
        console.log('da chay roi day')
    }
</script>


                     <!-- Phan up anh -->
            <div class="form-group-2">
                <label for="fullname" class="form-label">Meta-SEO:</label>
                <div class="ver-covered-input">
                    <span class="checkbox-text">Từ khóa: </span>
                    <!--<input id="meta_keyword" name="meta_keyword" type="text" placeholder="" class="form-control">-->
                    <input name="meta_keyword" id="meta_keyword" class="form-control" value="<%=meta_keyword %>" />
                    <span class="checkbox-text">Mô tả: </span>
                    <!--<input id="fullname" name="fullname" type="text" placeholder="" class="form-control">-->
                    <input name="meta_desc" id="meta_desc" class="form-control" value="<%=meta_desc %>" />
                </div>

            </div>
            <div class="form-group-2">
                <label for="fullname" class="form-label">Phê duyệt:</label>
                <div class="ver-covered-input">
                <div class="sub-hor-covered-input">

                    <div class="four-divided">
                        <label for="StatusId" class="checkbox-text">Trạng thái đăng tin:</label>
                        <select name="StatusId" id="StatusId" class="form-control">
                                    <option value="0">--Lựa chọn--</option>
                                    <option value="4">Đăng</option>
                                    <option value="2">Lưu</option>
                                </select>
                    </div>
                    <div class="four-divided">
                        <label for="DateCreater" class="checkbox-text">Ngày đăng:</label>
                        <input id="DateCreater" name="DateCreater" value="<%=DateNow %>"  class="form-control">
                    </div>
                    
                </div>
                <div class="checkbox-div"><input type="checkbox" checked="checked" name="attach_product" id="attach_product" value="1"
                    class="ckeckbox-children"><span class="checkbox-text">Tiếp tục gửi tin để thiết lập</span></div>
                    <div class="centered-item">
                        <!--<button class="form-submit">Đăng Ký</button>-->
                        <input class="form-submit" type="button" name="Button" value="<%=btnName %>" onclick="javascript: SendToOneCat();" />
                        <!--<button class="form-submit">Hủy</button>-->
                        <input class="form-submit" type="reset" name="Button" value="Hủy"  />
                            <%
                                    if iStatus  = "edit" then 
                                        %>
                                        <input type="hidden" name="NewsId" value="<%=NewsId%>">
                                        <input type="hidden" name="iStatus" value="edit">
                                        <input type="Hidden" name="old_PictureId" value="<%=PictureId%>">
                                        <input type="hidden" name="old_Note" value="<%=Note%>">
                                        <input type="Hidden" name="sCatId" value="<%=CatId%>">


                                        <%  
                                    else
                                        %>
                                        <input type="hidden" name="iStatus" value="add">
                                        <input type="Hidden" name="old_PictureId" value="0">

                                        <%    
                                          
                                    end  if
                                     
                                        %>
                    </div>
            </div>
            </div>
        </form>
            </div>
        </div>
           </div>
</div>
        <%Call Footer()%>
        <script  type="text/javascript" src="../inc/news.js"></script>
</body>
</html>
<script  type="text/javascript">
    function checkIsNumber(field) {
        yn = isNaN(field.value)
        if (yn) {
            alert("Chỉ chấp nhận dữ liệu dạng số. \nVui lòng nhập lại");
            field.select();
            return;
        }
    }
    function disSize() {
        var strTemp = document.fInsert.f_Size.value;
        var lg = strTemp.length;
        var x = strTemp.charAt(lg - 1);
        if ((checkInteger(x) == 1) || (x == 'x') || (x == 'm')
			|| (x == 'M') || (x == 'c') || (x == 'C') || (x == 'X') || (x == '.') || (x == ',')) {
            document.fInsert.f_Size.value = strTemp;
        }
        else {
            strTemp = strTemp.substring(0, strTemp.length - 1);
            document.fInsert.f_Size.value = strTemp;
        }
    }
 
</script>
<script type="text/javascript">
    CKEDITOR.replace('bodyx', {
        width: '100%',
        height: '500px'
    });
    //CKEDITOR.replace('F_Tskt');
    //  CKEDITOR.replace('Fconfig');
</script>

<link href="/css/date-picker.css" rel="stylesheet" />
<script src="/script/jQuery-2.1.4.min.js"></script>
<script src="/script/date-picker.js"></script>
<link href="/css/date-picker.css" rel="stylesheet" />


<script type="text/javascript">
    function ksa() {
        $.ajax({
            type: "POST",
            url: "/ksa.asp",
            data: {},
            cache: false,
            dataType: "html",
            success: function (rs_) {
                //  alert(rs_)
                $("#keepSessionAlive").val(rs_);
            },
            error: function (rs_) {
                alert("Đã có lỗi sảy ra, vui lòng F5 trình duyệt hoặc gọi kỹ thuật viên.");
            },
        });
    }
    $(document).ready(function () {
        setInterval(ksa, 1000);
    });
      $.datetimepicker.setLocale('vi');
      jQuery(function () {
          jQuery('#DateCreater').datetimepicker({
              format: 'd-m-Y H:i',
              timepicker: true,
              hours12: false,
              lang: 'vi',
              yearStart: 2021,
          });
      });
    
</script>
<script  type="text/javascript">

	function DisMoneyThis(field)
	{
		GTotal	=	field.value	
		if(field.value!=0)
			field.value =DisMoney(GTotal);			

	}
	function checkIsNumber(field)
	{
		yn = isNaN(field.value)
		if (yn)
		{
			alert("Chỉ chấp nhận dữ liệu dạng số. \nVui lòng nhập lại");
			field.value = 0;
			field.select();
			return;
		}
	
	}
function GetMoneyText(MFrom,MObject)
{
	str = "iMoney = document."+MFrom+"."+MObject+".value;";
	eval(str);
	if(iMoney==''||iMoney==0) 
		iMoney	=	0
	else
	{	
		for(k=1;k<=5;k++)
		{
			iMoney = iMoney.replace(",","")
		}
		for(k=1;k<=5;k++)
		{
			iMoney = iMoney.replace(".","")
		}
		
		iMoney	=	parseInt(iMoney);
	}	
	return 	iMoney
}		
</script>