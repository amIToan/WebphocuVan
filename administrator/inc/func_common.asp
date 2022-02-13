<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->


<%
function Dis_str_money(moneys)
    if not isNumeric(moneys) then
        Dis_str_money = 0
        exit function
    end if
	moneys	=	round(Cdbl(moneys))
	str_dao=Daochuoi(moneys)
	lenn=Len(str_dao)
	k3=1
	money_re=""
	for bieni=1 to lenn
		money_re=money_re+Mid(str_dao,bieni,1)
		if k3=3 and bieni< lenn then
			money_re=money_re+"."
			k3=1
		else
			k3=k3+1
		end if
	next
	Dis_str_money=Daochuoi(money_re)
end function
function Daochuoi(str_money)
	dim str_temp
	k_lengh=len(str_money)
	for bieni=1 to k_lengh
		str_temp=MID(str_money,bieni,1)+str_temp
		'Response.write("i="& i &" str_money "& str_money)
	next
	Daochuoi=str_temp
end function
%>
<%Sub header()%>
<div style="width: 100%;  background:#001e33; border-bottom: solid 2px #EC3237; color:#FFF;" >
    <table style="width: 100%; margin: auto;">
        <tr>
            <td>
                <table style="width: 100%" border="0" class="CTxtContent">
                    <tr>
                        <td>
                            <span style="color:#FFF;">
                            <%if Session("staffimg") <> "" then%>
                               
                                <%else%>
                            
                                <%end if%>
                               
                                
                                    Online:
			        <%	
			        iSubTime	=	DateAdd("n",-60,Now)
			        sqlon="select Count(ID) as iCount from MOnline where TimeDays > '"& iSubTime &"'"
			        Set rstemp = Server.CreateObject("ADODB.Recordset")
			        on error Resume next
			        rstemp.open sqlon,Con,3
			        if not rstemp.eof then
				        Response.Write(rstemp("iCount"))
			        end if
			        set rstemp = nothing
                    %>
                                    <br />
                                    IP đăng nhập: <%=Request.ServerVariables("REMOTE_ADDR")%>
                                </span>
                        </td>
                        <th style="text-align:right; padding-right:10%;">
                           <span style="color:#FFF;">ADMIN  -  QUẢN TRỊ</span> 
                        </th>
                        
                       
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</div>
<%End sub%>

<%Sub Footer()%>
<div style="height: 50px;"></div>
<table style="width: 100%; border-top: solid 1px #808080;background-color:#E6E6E6;" class="w3-table w3-table-all">
    <tr>
        <td>
            <div class="CSubTitle w3-padding w3-center">
                &copy;<%=year(now())%> - <%=company%>
            </div>
        </td>
    </tr>
</table>
<%End sub%>

<%sub MenuVertical(status)%>
<script src="/javascript/jquery-2.2.2.min.js"></script>
    <ul class="ul">
        <li class="wp-cate">Quản Trị Hệ Thống
            
        </li>
        <li>
            <div class="wp-cate-Item">
                <img src="/images/icon/en.jpg" style="height: 22px; cursor: pointer;" onclick="changeLang('EN');" />
                <img src="/images/icon/vn.jpg" style="height: 22px; cursor: pointer;" onclick="changeLang('VN');" />
            </div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/system/cat_list.asp">Quản lý chuyên mục</a> </div>
        </li>

        <li>
            <div class="wp-cate-Item"><a href="/administrator/system/Company_edit.asp">Thông tin công ty</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/yahoo/Yahoo_list.asp">Hỗ trợ trực tuyến</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/user/user_viewprofile.asp">Thông tin tài khoản</a></div>
        </li>
        
        <li class="wp-cate">Quản Trị Nội Dung</li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/news/news_addedit.asp?iStatus=add">Bài viết mới</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/news/news_search.asp">Tìm & sửa nội dung</a></div>
        </li>
		<li>
            <div class="wp-cate-Item"><a href="/administrator/news/list_products.asp">Quản lý đơn hàng</a></div>
        </li>
		<li>
            <div class="wp-cate-Item"><a href="/up_sitemap.asp">Cập nhật sitemap</a></div>
        </li>
        <li style="display:none;">
            <div class="wp-cate-Item"><a href="/administrator/FAQ/faq_list.asp">Quản lý liên hệ</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/panel/paner_list.asp">Đối tác</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/FAQ/faq_list.asp">Quản lý FAQ</a></div>
            <div style="display:none;" class="wp-cate-Item w3-hide"><a href="/administrator/FAQ/ykien_list.asp">Quản lý ý kiến KH</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/thongke/sl.asp">Biên tập viên</a></div>
        </li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/thongke/sl_category.asp">Tin theo chuyên mục</a></div>
        </li>
         <li>
            <div class="wp-cate-Item"><a href="/administrator/tools/ads_list.asp">Tin quảng cáo</a></div>
        </li>        
        <li class="wp-cate">Quản Trị User</li>
        <li>
            <div class="wp-cate-Item"><a href="/administrator/system/user_list.asp">Danh sách User</a></div>
        </li>
   		<li class="wp-cate"> <div class="wp-cate-Item"><a href="/administrator">Thoát</a></div></li>
           
    </ul>
    <script type="text/javascript">
        function changeLang(str) {
            //alert(str);
            $.ajax({
                url: "/include/ajax_common.asp",
                type: 'POST',
                data: {
                    "_key": "changeLang",
                    "lang": str
                },
                cache: false,
                dataType: "html",
                success: function (rs) {
                    location.reload();
                },
                error: function () {
                    alert("đã có lỗi sảy ra");
                }
            });
        }
    </script>
<%end sub%>

<%
    Sub VerticalMenuFunction(NameTab,LinkTab,LinkArrayMenu,TitleArrayMenu)
%>
<script src="../../Scripts/jquery-1.8.2.min.js"></script>
<script src="../../Scripts/jquery.js"></script>
<script src="../../Scripts/jquery.easing.js"></script>
<link href="../../css/menu.css" rel="stylesheet" />
<script src="../../Scripts/script.js"></script>

<li>
    <a href="<%=LinkTab%>"><%=NameTab%></a>
    <%if Ubound(LinkArrayMenu) > 0 then%>
    <ul>
        <%
    for j =0 to Ubound(LinkArrayMenu)
 	    if Trim(LinkArrayMenu(j))<>"" then 
		    varLink1="/administrator/"+LinkArrayMenu(j)
	    else 
	  	    varLink1="welcome.asp"
	    End if 
        %>
        <li><a href="<%=varLink1%>"><%=TitleArrayMenu(j)%></a></li>
        <%
    next
        %>
    </ul>
    <%end if %>
</li>
<script type="text/javascript">

    function initMenu() {
        $('#menu ul').hide();
        $('#menu ul:first').show();
        $('#menu li a').click(
          function () {
              var checkElement = $(this).next();
              if ((!checkElement.is('ul')) && (checkElement.is(':visible'))) {
                  checkElement.slideUp('normal');
                  return false;
              }
              if ((checkElement.is('ul')) && (!checkElement.is(':visible'))) {
                  checkElement.slideDown('normal');
                  return false;
              }
          }
          );
    }
    $(document).ready(function () { initMenu(); });
</script>
<%
    end sub
%>

<%Sub menuOld()
	tab_menu	=	"../../images/tab_menu.png"
%>
<table style="width: 100%;" border="0">
    <tr>
        <td background="../../images/TopMenu.jpg">
            <table cellpadding="0" cellspacing="0" align="center" width="1000px">
                <tr>
                    <td align="right">
                        <table height="40px" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <%
          
		NameTab			=	"Cập nhật"
		iTab = 1 
		TitleArrayMenu	=	array("Trang chủ","Thông báo","Thông tin công ty","Quản lý User","Quản lý chuyên mục","Hỗ trợ yahoo","Tiêu biểu theo chủ đề","Cập nhật tỷ giá", "Cập nhật bảng giá","Nhập liệu nhà cung cấp","Bảng màu sản phẩm","Quản lý ảnh","Tìm kiếm ảnh")
		LinkArrayMenu		=	array("welcome.asp","system/upThongBao.asp","system/Company_edit.asp","system/user_list.asp","system/cat_list.asp","yahoo/Yahoo_list.asp","news/TopicTypical.asp","system/exchange_list.asp","system/price_list.asp","Provider/Provider_list.asp","tools/color_table.asp","picture/picture_list.asp","picture/picture_Search.asp")
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)	

		NameTab			=	"Nhập liệu"
		iTab = iTab+1
		strTitleArrayMenu	=	"Nhập tin mới"
		strLinkArrayMenu	=	"news/news_addedit.asp?iStatus=add"
		strTitleArrayMenu	=	strTitleArrayMenu+";Tìm và sửa tất cả"
		strLinkArrayMenu	=	strLinkArrayMenu+";news/news_search.asp"		
		
		TitleArrayMenu		=	split(strTitleArrayMenu,";")
		LinkArrayMenu		=	split(strLinkArrayMenu,";")
		
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)	

		NameTab			=	"QLý hàng"
		iTab = iTab+1
		TitleArrayMenu	=	array("Quản lý đơn hàng","Thống kê phần mềm","Đặt hàng theo yêu cầu","Quản lý website","Trả nhà cung cấp")
		LinkArrayMenu		=	array("donhang/order_list.asp","donhang/order_detail.asp","donhang/OrderContact.asp","xseotitle/cus_information.asp","TraSach/TraSach.asp")
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)			  

		NameTab			=	"Khách hàng"
		iTab = iTab+1
		TitleArrayMenu	=	array("Ý kiến khách hàng","Câu hỏi thường gặp","Thông tin","Giao dịch đơn hàng")
		LinkArrayMenu		=	array("FAQ/ykien_list.asp","FAQ/faq_list.asp","customer/customer_list.asp","customer/History_list.asp")
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)	
		  	
		NameTab			=	"Công cụ"
		iTab = iTab+1
		TitleArrayMenu	=	array("Banner và quảng cáo","Quản trị khuyến mãi","Quảng cáo qua email")
		LinkArrayMenu		=	array("tools/ads_list.asp","KhuyenMai/DanhSachKM.asp","customer/List_email.asp")
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)	

		NameTab			=	"Nhân sự"
		iTab = iTab+1
		TitleArrayMenu	=	array("Hồ sơ","Phòng ban","Chức danh","Quản trị công việc","Quản lý tiền lương")
		LinkArrayMenu		=	array("XHR/stafflist.asp","XHR/PhongBanList.asp","XHR/ChucDanhList.asp","XHR/qlNhansu.asp","XHR/qlTienLuong.asp")
		call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)

		NameTab			=	"Thống kê"
		iTab = iTab+1
		TitleArrayMenu	=	array("Biên tập viên","Tin theo chuyên mục","Thống kê danh sách")
		LinkArrayMenu		=	array("thongke/sl.asp","thongke/sl_category.asp","thongke/ReportListBook.asp")
        call MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)
							  
                                %>
                            </tr>
                        </table>
                    </td>
                    <td align="right">
                        <a href="/administrator/user/user_viewprofile.asp">
                            <img src="../../images/icons/icon_profile.gif" alt="Thông tin" width="15" height="15" border="0" />
                        </a>&nbsp;&nbsp; <a href="/administrator/">
                            <img src="../../images/icons/icon_go_right.gif" alt="logout" width="15" height="15" border="0" /></a>&nbsp;&nbsp;<a href="javascript: winpopup('../../include/Guide.asp','all',700,500);" style="cursor: help;"><img src="../../images/icons/Qm.gif" width="16" height="16" border="0" alt="Hướng dẫn chung"></a>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td align="center" class="CTxtContent"><%Call TitlePage(Title_This_Page,img)%></td>
    </tr>
</table>
<script language="javascript">
    function change_bg_menu(stt, path_image) {
        document.getElementById("mainmenu" + stt).style.background = "url(" + path_image + ")";
        if (path_image == '')
            document.getElementById("font_menu_top" + stt).color = '#055F91';
        else
            document.getElementById("font_menu_top" + stt).color = '#FFFFFF';

    }
</script>
<%End sub%>
<%
sub MenuAdminTab(iTab,NameTab,tab_menu,LinkArrayMenu,TitleArrayMenu)
%>
<td id="mainmenu<%=iTab%>" style="background-repeat: no-repeat; background-position: center;" align="center" width="91px" onmouseover="javascript:change_bg_menu(<%=iTab%>,'<%=tab_menu%>');" onmouseout="javascript:change_bg_menu(<%=iTab%>,'');">
    <a href="#"><font id="font_menu_top<%=iTab%>" class="CTxtTitle10"><%=NameTab%></font></a>
    <%if Ubound(LinkArrayMenu) <> 0 then%>
    <script type="text/javascript">
        jkmegamenu.definemenu("mainmenu<%=iTab%>", "altmainmenu<%=iTab%>")
    </script>
    <div id="altmainmenu<%=iTab%>" class="Cmainmenutab">
        <table border="0" cellspacing="0" cellpadding="0" onmouseover="javascript:change_bg_menu(<%=iTab%>,'<%=tab_menu%>');" onmouseout="javascript:change_bg_menu(<%=iTab%>,'');">
            <tr>
                <td height="5" background="../../images/BGblue.png"></td>
            </tr>
            <%
for j =0 to Ubound(LinkArrayMenu)
 	if Trim(LinkArrayMenu(j))<>"" then 
		varLink1="/administrator/"+LinkArrayMenu(j)
	else 
	  	varLink1="welcome.asp"
	End if 
            %>
            <tr>
                <td height="35" id="mainmenu<%=iTab%><%=j%>" onmouseover="javascript:change_bg_menu(<%=iTab%><%=j%>,'../../images/BGblue.png');" onmouseout="javascript:change_bg_menu(<%=iTab%><%=j%>,'');"><a href="<%=varLink1%>"><font id="font_menu_top<%=iTab%><%=j%>" color="#055F91" class="CTxtTitle10"> &nbsp; + <%=TitleArrayMenu(j)%> &nbsp;</font></a></td>
            </tr>
            <%
next
            %>
        </table>
    </div>
    <%end if%>  
</td>
<%end sub%>


<%Sub Titlepage(Title_This_Page,img)%>
<%
if img ="" then
	img	="../../images/icons/icon_report.gif"	
end if
%>
<table width="998" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
        <td>
            <table width="998" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="85%" class="CTieuDeNho">&nbsp;&nbsp;&nbsp;<img src="<%=img%>" height="48" border="0" align="absmiddle" />&nbsp;&nbsp;<%=Title_This_Page%>
                    </td>
                    <td width="15%" align="right"><font size="1" face="verdana"><strong><font color="#999999">
		  	Hôm nay, ngày <%=Day(ConvertTime(now))%>/<%=Month(ConvertTime(now))%>/<%=Year(ConvertTime(now))%>
		  </font></strong></font></td>
                </tr>
            </table>
            <table width="998" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td bgcolor="1079E7">
                        <img src="/administrator/images/1x1.gif" width="1" height="1"></td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%End sub%>
<%Function Authenticate(ListOfCategory)
	'ListOfCategory=None: Chỉ kiểm tra có tồn tại user đăng nhập
	'				Admin: Kiểm tra quyền admin mới được vào
	'				CategoryId: Chỉ những user có quyền vào chuyên mục này mới có thể vào được
	Select case ListOfCategory
		case "None"
			if Trim(session("user"))="" then
				response.Redirect("/administrator/default.asp")
				response.End()
			end if
		case "Admin"
			if Trim(session("user"))="" then
				response.Redirect("/administrator/default.asp")
				response.End()
			elseif Session("LstCat")<>"0" or Session("LstRole")<>"0ad" then
				response.Redirect("/administrator/info.asp") 'Trả về trang hiện ra thông báo kô có quyền vào vùng đó
				response.End()
			end if

		case "QuanLy"
			if Trim(session("user"))="" then
				response.Redirect("/administrator/default.asp")
				response.End()
			elseif Session("LstCat")<>"0" or Session("LstRole")<>"0ad" then
				response.Redirect("/administrator/info.asp") 'Trả về trang hiện ra thông báo kô có quyền vào vùng đó
				response.End()
			end if

		case else
			if Trim(session("user"))="" then
				response.Redirect("/administrator/default.asp")
				response.End()
			elseif CompareRoleCat(session("LstCat"),ListOfCategory)=0 then
				response.Redirect("/administrator/info.asp") 'Trả về trang hiện ra thông báo kô có quyền vào vùng đó
				response.End()
			end if
	End select
End Function%>



<%Function AuthenticateWithRole(CategoryId,LstRole,Role)
	'CategoryId: Chuyên mục cần xử lý
	'LstRole: Danh sách quyền truyền vào
	'Role: 	=NONE
	'		=ed: CategoryId tương ứng trong LstRole phải Role >= ed
	'		=se: CategoryId tương ứng trong LstRole phải Role >= se
	'		=ap: CategoryId tương ứng trong LstRole phải Role >= ap
	'		=ad: CategoryId tương ứng trong LstRole phải Role >= ad
	LstRole	= "0ad"' tuannv
	if LstRole="" then
		response.Redirect("/administrator/default.asp")
		response.End()
		Exit Function
	end if
	
	if LstRole="0ed" or  LstRole="0se" or  LstRole="0ap" or  LstRole="0ad" then
		if CompareRole(Role,Right(LstRole,2))>0 then
			response.Redirect("/administrator/info.asp")
			response.end()
		end if
	else
		strRole=GetRoleOfCat_FromListRole(CategoryId,LstRole)
		'response.Write(CatId & ":" & LstRole)
		'response.End()
		if CompareRole(Role,strRole)>0 or strRole="" then
			response.Redirect("/administrator/info.asp")
			response.end()
		end if
	end if
End Function%>

<%Function CompareRoleCat(ListOfCat,Cat)
	'Resulf <>0: Cat is in ListOfCat; else not in ListOfCat
	CompareRoleCat=Instr(" " & Trim(ListOfCat) & " "," " & Trim(Cat) & " ")
End Function%>

<%Function ConvertTime(thedate)
	'Convert time when hosting server do not match time-belt with customer
	ConvertTime=DateAdd("h",0,thedate)
End Function%>

<%Function GetNameOfLanguage(LanguageId)
	sql="select LanguageName from Language where LanguageId='" & LanguageId & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		GetNameOfLanguage=""
		exit function
	end if
	GetNameOfLanguage=Trim(rs("LanguageName"))
	rs.close
	set rs=nothing
End Function%>

<%Function GetNameOfCategory(CatId)
	sql="select CategoryName from NewsCategory where CategoryId=" & CatId
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		GetNameOfCategory=""
		exit function
	end if
	GetNameOfCategory=Trim(rs("CategoryName"))
	rs.close
	set rs=nothing
End Function%>

<%Sub ListStatusOfCategory(StatusSelect)%>
<select name="CategoryStatus" class="form-control" style="width:200px;">
    <option value="0" <%if StatusSelect=0 then%> selected<%End if%>>Tắt chức năng</option>
    <option value="1" <%if StatusSelect=1 then%> selected<%End if%>>Menu dọc</option>
    <option value="5" <%if StatusSelect=5 then%> selected<%End if%>>Menu dọc + Nội dung</option>
    <option value="2" <%if StatusSelect=2 then%> selected<%End if%>>Menu ngang</option>
    <option value="4" <%if StatusSelect=4 then%> selected<%End if%>>Chuyên đề</option>
    <option value="3" <%if StatusSelect=3 then%> selected<%End if%>>Quản lý</option>
</select>
<%End Sub%>

<%Function GetNameOfCategoryStatus(CatStatus)
	Select case CatStatus
		case 0
			GetNameOfCategoryStatus="Tắt chức năng"
		case 1
			GetNameOfCategoryStatus="Menu dọc"
        case 5
			GetNameOfCategoryStatus="Menu dọc + Nội dung"
		case 2
			GetNameOfCategoryStatus="Menu ngang"
		case 3
			GetNameOfCategoryStatus="Quản lý"
		case 4
			GetNameOfCategoryStatus="Chuyên đề"	
       	
	End Select
End Function%>

<%Sub List_Language(LanguageSelect)
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="select LanguageId,LanguageName from Language"
	rs.Open sql, con, 1
	
	response.Write "<select name=""languageid"" id=""languageid"">"
    
	Do while not rs.eof
		response.Write "<option value=""" & rs("LanguageId")  & """"
		if rs("LanguageId")=LanguageSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(rs("LanguageName") & "</option>")
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
	set rs=nothing
End sub%>

<%Sub List_Event(EventSelect,EventTitle,EventCount)
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="select top " & EventCount & " EventId,EventName from Event order by EventId desc"
	rs.Open sql, con, 1
	
	response.Write "<select name=""eventid"" id=""eventid"">"
    	response.Write("<option value=""0"">" & EventTitle & "</option>")
	Do while not rs.eof
		response.Write "<option value=""" & rs("eventid")  & """"
		if rs("eventid")=EventSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(rs("EventName") & "</option>")
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
	set rs=nothing
End sub%>

<%Sub phantrang(page,pagecount,pageperbook)%>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><u>Trang:</u></strong> 
          	<%
          	minpage=page-Clng((pageperbook-1)/2)
          	maxpage=page+Clng((pageperbook-1)/2)
			'response.write "min=" & minpage & ",max=" & maxpage
          	if minpage<1 and maxpage>pagecount then
          	'Vuot qua' ca can dau va` can cuoi
          		minpage=1
          		maxpage=pagecount
          	elseif minpage<1 then
          	'Truong hop chi co' minpage nho hon can duoi
          		minpage=1
          		maxpage=minpage+pageperbook-1
          		if maxpage>pagecount then
          			maxpage=pagecount
          		end if
          	elseif maxpage>pagecount then
          	'Truong hop chi co' maxpage lon hon can tren
          		maxpage=pagecount
          		minpage=maxpage-pageperbook+1
          		if minpage<1 then
          			minpage=1
          		end if
          	end if
          	
			'Xu ly' URL QUERY_STRING
          	if request.ServerVariables("QUERY_STRING")<>"" then
				Dim ArrURLa,ArrURLb
				ArrURLa=Split(request.ServerVariables("QUERY_STRING"),"&")
				ArrURLb=Filter(ArrURLa,"page=",false)
				sURL=Trim(Join(ArrURLb,"&"))
				if sURL<>"" then
					sURL=request.ServerVariables("SCRIPT_NAME") & "?" & sURL & "&"
				else
					sURL=request.ServerVariables("SCRIPT_NAME") & "?"
				end if
			else
				sURL=request.ServerVariables("SCRIPT_NAME") & "?"
			end if
			
          	for i=minpage to maxpage
				if i<>page then
					response.Write "<a href=""" & sURL & "page=" & i & """>" & i & "</a>&nbsp;|&nbsp;"
				else
					response.Write "<font color=""red""><b>" & i & "</b></font>&nbsp;|&nbsp;"
					NEXT_PAGE=i+1
				end if
			Next
			
			if NEXT_PAGE<=pagecount then
				response.Write "<a href=""" & sURL & "page=" & NEXT_PAGE & """ style=""text-decoration: none"">Ti&#7871;p&#8250;&#8250;</a>"
			end if%>
	</font>
<%End sub%>
<%Function CheckUserExist(Username)
	'Kiểm tra xem có tồn tại User này hay không.
	'Kết quả trả về: 
	'				0: Không tồn tại User
	'				<>0: Tồn tại User 
	
	if LCase(Username)="admin" then
		CheckUserExist=1
		exit Function
	end if
	sql="select count(Username) as dem from [User] where Username=N'" & username  & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	CheckUserExist=Clng(rs("dem"))
	rs.close
	set rs=nothing
End Function%>

<%Function CheckUserRoleExist(Username)
	'Kiểm tra xem có tồn tại UserRole này hay không.
	'Kết quả trả về: 
	'				0: Không có quyền
	'				<>0: Có quyền

	sql="select count(Username) as dem from Userdistribution where Username=N'" & username  & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		CheckUserRoleExist=Clng(rs("dem"))
	rs.close
	set rs=nothing
End Function%>

<%Function RandomPassword(myLength)
	'These constant are the minimum and maximum length for random
	'length passwords.  Adjust these values to your needs.
	Const minLength = 6
	Const maxLength = 20
	
	Dim X, Y, strPW
	
	If myLength = 0 Then
		Randomize
		myLength = Int((maxLength * Rnd) + minLength)
	End If

	
	For X = 1 To myLength
		'Randomize the type of this character
		Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
		
		Select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))

		End Select
	Next
	
	RandomPassword = strPW
End Function%>

<%Sub Display_Images_Library(CatId,Page, FieldOrder,TypeOrder,ChooseToInsert)
	'OrderByField: PictureCaption, PictureId (Creationdate)
	'OrderByType: desc or none
	'ChooseToInsert=1: Hiển thị chức năng cho phép chọn ảnh để chèn vào bài viết
	'			   =0: Tắt chức năng cho phép chọn ảnh để chèn vào bài viết
%>
<script language="javascript">
    function onButtonClick(anhnho, anhto, border) {
        if (anhto != "") {
            imageTag = "<a href=\"javascript:winpopup('/administrator/picture/picture_view.asp','large_9.jpg',100,200);\">";
            imageTag += "<IMG src=\"" + anhnho + "\" ";
        }
        else {
            imageTag = "<IMG src=\"" + anhnho + "\" ";
        }

        imageTag += "alt=\"Ảnh minh họa\" ";
        imageTag += "align=\"center\" ";
        imageTag += "border=\"" + border + "\">";
        if (anhto != "") {
            imageTag += "</a>";
        }
        opener.InsertNewImage(imageTag);
        window.close();
    }
    function onButtonClick_path(anhnho, anhto, ImagePath, border) {
        if (anhto != "") {
            imageTag = "<a href=\"javascript:openImage('" + ImagePath + anhto + "');\">";
            imageTag += "<IMG src=\"" + ImagePath + anhnho + "\" ";
        }
        else {
            imageTag = "<IMG src=\"" + ImagePath + anhnho + "\" ";
        }

        imageTag += "alt=\"Ảnh minh họa\" ";
        imageTag += "align=\"center\" ";
        imageTag += "border=\"" + border + "\">";
        if (anhto != "") {
            imageTag += "</a>";
        }
        opener.InsertNewImage(imageTag);
        window.close();
    }
</script>
<%
	Select case FieldOrder
		case 0
			OrderByField="PictureId"
		case 1
			OrderByField="PictureCaption"
	End Select
	Select case TypeOrder
		case 0
			OrderByType="desc"
		case 1
			OrderByType=""
	End Select
	sql="SELECT *"
	sql=sql & " FROM Picture"
	sql=sql & " WHERE CategoryId=" & CatId
	sql=sql & " Order by " & OrderByField & " " & OrderByType
	
	'response.Write(sql)
	set rs=server.CreateObject("ADODB.Recordset")
	PAGE_PER_BOOK=5
	rs.PageSize = 9
	rs.open sql,con,1
	
	if rs.eof then
		rs.close
		set rs=nothing
		exit sub
	end if
	
	rs.AbsolutePage = CLng(page)
	i=0
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
    <%Do while not rs.eof and i<rs.pagesize
        j=0
    %>
    <tr>
        <%Do while not rs.eof and j<3%>
        <td width="33%" valign="top">
            <table align="center" cellpadding="2" cellspacing="2">
                <tr>
                    <td align="center" valign="top" colspan="3">
                        <img src="<%=NewsImagePath%><%=rs("SmallPictureFileName")%>" border="0" width="150"></td>
                </tr>
                <tr>
                    <td align="center" colspan="3"><font size="1" face="verdana"><%=rs("PictureCaption")%></font></td>
                </tr>
                <tr>
                    <td colspan="3" align="left"><font size="1" face="Arial, Helvetica, sans-serif">Tạo 
                  bởi: <%=rs("Creator")%> (<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>)<br>
				  <%if rs("IsHomePicture") then%>
	                  <img src="../images/update.gif" width="16" height="16" align="absmiddle"> Ảnh trang chủ<br>
				  <%End if%>
				  <%if rs("IsCatHomePicture") then%>
    	              <img src="../images/update.gif" width="16" height="16" align="absmiddle">Ảnh chuyên mục<br>
				  <%End if%>
				  <%if rs("statusId")="ap" or rs("statusId")="ad" then%>
	                  Trạng thái: Đã duyệt<br>
    	              Duyệt bởi: <%=rs("Approver")%> (<%=Day(ConvertTime(rs("ApproverDate")))%>/<%=Month(ConvertTime(rs("ApproverDate")))%>/<%=Year(ConvertTime(rs("ApproverDate")))%>)
				  <%end if%>
				</font></td>
                </tr>
                <tr>
                    <td align="center">
                        <%if rs("LargePictureFileName")<>"" then%>
                        <a href="javascript: winpopup('/administrator/picture/picture_view.asp','<%=rs("LargePictureFileName")%>',100,200);">Xem</a>
                        <%else
						response.Write("&nbsp;")
					end if%>				</td>
                    <%RolePower=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
				if RolePower="ap" or RolePower="ad" then%>
                    <td align="center">
                        <a href="javascript: winpopup('/administrator/picture/picture_edit.asp','<%=rs("pictureid")%>',420,300);">Sửa</a></td>
                    <%if Session("iQuanTri") = 1 then %>
                    <td align="center"><a href="javascript: winpopup('/administrator/picture/picture_delete.asp','<%=rs("pictureid")%>',300,150);">Xóa</a>
                        <%end if%>									</td>
                    <%
					Else
                    %>
                    <%end if%>
                    <td align="center"><font size="1" face="Arial, Helvetica, sans-serif">
					<%if ChooseToInsert=1 then%>
						<a href="javascript: onButtonClick('<%=NewsImagePath%><%=rs("SmallPictureFileName")%>','<%=rs("LargePictureFileName")%>','0');">Chọn</a>
					<%else%>&nbsp;
					<%end if%>
				</font></td>
                </tr>
                <tr>
                    <td colspan="3" align="center"><a href="javascript: onButtonClick_path('<%=rs("SmallPictureFileName")%>','<%=rs("LargePictureFileName")%>','<%=NewsImagePath%>','0');" style="text-decoration: none">Chèn</a>	</td>
                    <td align="center">&nbsp;</td>
                </tr>
            </table>
        </td>
        <%j=j+1
          	i=i+1
          	if j<3 then
          		rs.movenext
          	end if
          	Loop%>
    </tr>
    <%i=i+1
			if not rs.eof then
				rs.movenext
			end if
		Loop
    %>
    <tr>
        <td colspan="3" align="right">
            <%Call phantrang(page,rs.pagecount,PAGE_PER_BOOK)
			rs.close
			set rs=nothing
            %>
        </td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
</table>
<%End sub%>

<%Function GetMaxId(TableName, FieldNameId, sCondition)
	Dim Max,rsMaxId
	set rsMaxId=server.CreateObject("ADODB.Recordset")
	sql="select Max(" & FieldNameId & ") as MaxId from " & TableName
	if sCondition<>"" then
		sql=sql & " where sCondition"
	end if
	rsMaxId.Open sql, con, 1
	if IsNull(rsMaxId("MaxId")) then
		Max=1
	else
		Max=CLng(rsMaxId("MaxId")) +1
	end if
	rsMaxId.close
	set rsMaxId=nothing
	GetMaxId=Max
End Function%>

<%Sub Display_News_Event(CatId,Page, FieldOrder,TypeOrder)
	'OrderByField: EventName, EventId (Creationdate)
	'OrderByType: desc or none

	Select case FieldOrder
		case 0
			OrderByField="EventId"
		case 1
			OrderByField="EventName"
	End Select
	Select case TypeOrder
		case 0
			OrderByType="desc"
		case 1
			OrderByType=""
	End Select
	sql="SELECT *"
	sql=sql & " FROM Event"
	sql=sql & " WHERE CategoryId=" & CatId
	sql=sql & " Order by " & OrderByField & " " & OrderByType
	
	'response.Write(sql)
	set rs=server.CreateObject("ADODB.Recordset")
	PAGE_PER_BOOK=5
	rs.PageSize = 20
	rs.open sql,con,1
	
	if rs.eof then
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.AbsolutePage = CLng(page)
	i=0
	j=(page-1) * rs.PageSize+1
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
    <%Do while not rs.eof and i<rs.pagesize%>
    <tr>
        <td align="right" valign="top">
            <%if Trim(rs("EventImages"))<>"" then%>
            <img src="<%=NewsImagePath%><%=rs("EventImages")%>">
            <%end if%>
        </td>
        <td valign="top">
            <%if rs("IsHomeEvent") then%>
            <img src="../images/icon-affiliate.gif" width="16" height="16" border="0" align="absmiddle" alt="Tin ảnh của trang chủ">
            <%end if%>
            <%if rs("IsCatHomeEvent") then%>
            <img src="../images/icon-campaign.gif" width="16" height="16" border="0" align="absmiddle" alt="Tin ảnh của chuyên mục">
            <%end if%>
            <font size="2" face="Arial, Helvetica, sans-serif"><strong><a href="javascript: winpopup('event_viewcontent.asp','<%=rs("EventId")%>&CatId=<%=rs("CategoryId")%>',700,450);"><%=rs("EventName")%></a></strong></font>
            <br>
            <font size="1" face="Arial, Helvetica, sans-serif">
            (<em>Tạo: <%=rs("Creator")%>-<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>,
            	Sửa: <%=rs("Approver")%>-<%=Day(ConvertTime(rs("ApproverDate")))%>/<%=Month(ConvertTime(rs("ApproverDate")))%>/<%=Year(ConvertTime(rs("ApproverDate")))%>,
            Ngôn ngữ: <%=rs("LanguageId")%>
            </em>)</font></td>
        <td align="right" valign="top">
            <a href="javascript: winpopup('event_edit.asp','<%=rs("EventId")%>',420,300);">
                <img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"></a>
            <a href="javascript: winpopup('event_delete.asp','<%=rs("EventId")%>',300,200);">
                <img src="../images/icon_closed_topic.gif" width="15" height="15" border="0" align="absmiddle"></a></td>
    </tr>
    <%j=j+1
          i=i+1
          rs.movenext
       Loop%>
    <tr>
        <td colspan="3" align="right">
            <%Call phantrang(page,rs.pagecount,PAGE_PER_BOOK)
			rs.close
			set rs=nothing
            %>
        </td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
</table>
<%End sub%>
<%Sub ShowPicture(PictureId,NoteAvailable,PictureAlign)
	'NoteAvailable=0: No Display Note
	'NoteAvailable=1: Display Note
	dim rsPic
	set rsPic=Server.CreateObject("ADODB.Recordset")
	sql="SELECT SmallPictureFileName,LargePictureFileName,PictureCaption,PictureAuthor"
	sql=sql & " FROM Picture"
	sql=sql & " WHERE PictureId=" & PictureId
	
	rsPic.open sql,con,1
	if rsPic.eof then
		rsPic.close
		set rsPic=nothing
		exit sub
	end if
%>
<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000080" width="100" id="AutoNumber49" align="<%=PictureAlign%>">
    <tr>
        <td width="1%">
            <%if rsPic("LargePictureFileName")<>"" then%>
            <a href="javascript: winpopup('/administrator/picture/picture_view.asp','<%=rsPic("LargePictureFileName")%>','0','0');">
                <img src="<%=NewsImagePath%><%=rsPic("SmallPictureFileName")%>" border="0">
            </a>
            <%Else%>
            <img src="<%=NewsImagePath%><%=rsPic("SmallPictureFileName")%>" border="0">
            <%End if%>
        </td>
    </tr>
    <%if NoteAvailable=1 then%>
    <tr>
        <td align="center"><font face="Arial" size="1"><%=rsPic("PictureCaption")%><%if Trim(rsPic("PictureAuthor"))<>"" then%>&nbsp;Ảnh: <%=rsPic("PictureAuthor")%><%End if%></font></td>
    </tr>
    <%end if%>
</table>
<%rsPic.close
set rsPic=nothing
End Sub%>
<%Function CheckUserProcess(User,CatId,NewsId)
	'User: co' 3 gia' tri Editor,GroupSenior,Approver
	'Chức năng: Kiểm tra xem User này có tham gia vào quá trình nhập tin hay không?
	'VD: Editor=bientapvien: Có nghĩa là User bientapvien tham gia va`o qua' trình nhập tin 
	'ở giai đoạn Editor
	'Nếu = NULL: Không tham gia
	'	<> NULL: tham gia
	sql="Select n." & User & " as btv from NewsDistribution d,News n where d.CategoryId=" & CatId & " and n.NewsId=" & Newsid & " and n.NewsId=d.NewsId"

	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		if rs.eof then
			CheckUserProcess=NULL
		else
			CheckUserProcess=rs("btv")
		end if
	rs.close
	set rs=nothing
End Function%>
<%Function CheckStatusWithCategoryId(NewsId,CatId,StatusId,LstRole)
	'StatusId: 	1: Đánh dấu
	'			2: Gửi trở lại cấp dưới yêu cầu sửa
	'			3: Gửi trở lên cấp trên duyệt
	'			4: Gửi lên mạng
	Dim strRole
	strRole=""
	strRole=GetRoleOfCat_FromListRole(CatId,LstRole)
	Select case StatusId
		case 1
			CheckStatusWithCategoryId=strRole & "ma"
			exit Function
		case 2
			if strRole="ed" then
			'Nếu quyền là Biên tập viên thì không có cấp dưới nào để gửi
				CheckStatusWithCategoryId="B&#7841;n kh&#244;ng c&#243; c&#7845;p d&#432;&#7899;i n&#224;o."
			else
				Select case strRole
					case "se"
						CheckName="Editor"
					case "ap"
						CheckName="GroupSenior"
					case "ad"
						CheckName="Approver"
				End select
				if IsNull(CheckUserProcess(CheckName,CatId,NewsId)) then
					CheckStatusWithCategoryId="Tin n&#224;y kh&#244;ng c&#243; c&#7845;p d&#432;&#7899;i n&#224;o g&#7917;i l&#234;n"
				else
					Select case strRole
						case "se"
							CheckStatusWithCategoryId="seed"
						case "ap"
							CheckStatusWithCategoryId="apse"
						case "ad"
							CheckStatusWithCategoryId="adap"
					End select
				end if
			end if
		case 3
			if strRole="ad" then
			'Nếu là quyền Administrator thì không có quyền gửi lên cấp trên duyệt
				CheckStatusWithCategoryId="B&#7841;n kh&#244;ng c&#243; c&#7845;p tr&#234;n n&#224;o &#273;&#7875; g&#7917;i l&#234;n"
			else
				Select case strRole
					case "ed"
						CheckStatusWithCategoryId="edse"
					case "se"
						CheckStatusWithCategoryId="seap"
					case "ap"
						CheckStatusWithCategoryId="apad"
				End select
			end if
		case 4
			if strRole<>"ap" and strRole<>"ad" then
				CheckStatusWithCategoryId="B&#7841;n kh&#244;ng c&#243; quy&#7873;n g&#7917;i tin l&#234;n m&#7841;ng &#7903; chuy&#234;n m&#7909;c n&#224;y"
			elseif strRole="ap" then
				CheckStatusWithCategoryId="apap"
			elseif strRole="ad" then
				CheckStatusWithCategoryId="adad"
			end if
	End select
End Function%>

<%Function GetFirstCategoryId_With_AP_Role(LstRole)
'Lấy chuyên mục đầu tiên có quyền ap hoặc ad trong chuỗi LstRole truyền vào
 
 		pos=Instr(LstRole,"ap") 'Vị trí na`y là vị trí của chữ "a" trong ap vừa tìm
		if pos=0 then
			pos=Instr(LstRole,"ad")
		end if
		if pos=0 then
			GetFirstCategoryId_With_AP_Role=1
			Exit Function 
		end if
		
		i=0
		s=""
		ch=""
		Do while i=0
			pos=pos-1
			if pos>0 then
				ch=Mid(LstRole,pos,1)
			else 
				ch=" "
			end if
			
			if ch=" " then
				i=1
			else
				s=ch & s
			end if
		Loop

		if Clng(s)=0 then 
			GetFirstCategoryId_With_AP_Role=1
		else
			GetFirstCategoryId_With_AP_Role=Clng(s)
		end if
End Function%>
<%Sub Display_Vote(CatId,Page, FieldOrder,TypeOrder)
	'OrderByField: EventName, EventId (Creationdate)
	'OrderByType: desc or none

	Select case FieldOrder
		case 0
			OrderByField="VoteID"
		case 1
			OrderByField="VoteTitle"
	End Select
	Select case TypeOrder
		case 0
			OrderByType="desc"
		case 1
			OrderByType=""
	End Select
	sql="SELECT *"
	sql=sql & " FROM Vote"
	sql=sql & " WHERE CategoryId=" & CatId
	sql=sql & " Order by " & OrderByField & " " & OrderByType
	
	set rs=server.CreateObject("ADODB.Recordset")
	PAGE_PER_BOOK=5
	rs.PageSize = 20
	rs.open sql,con,1
	
	if rs.eof then
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.AbsolutePage = CLng(page)
	i=0
	j=(page-1) * rs.PageSize+1
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
    <%Do while not rs.eof and i<rs.pagesize%>
    <tr>
        <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=j%>.</font></td>
        <td valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><strong><a href="javascript: winpopup('vote_view.asp','<%=rs("VoteId")%>&catid=<%=rs("CategoryId")%>',420,300);"><%=rs("VoteTitle")%></a></strong></font>
            <br>
            <font size="1" face="Arial, Helvetica, sans-serif">
            (<em>Tạo: <%=rs("Creator")%>-<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>,
            	Sửa lần cuối: <%=rs("Approver")%>-<%=Day(ConvertTime(rs("ApproverDate")))%>/<%=Month(ConvertTime(rs("ApproverDate")))%>/<%=Year(ConvertTime(rs("ApproverDate")))%>,
            <%if rs("statusId")="ap" or rs("statusId")="ad" then%>
            	<%if rs("IsHomeVote") then%>
            		<img src="../images/update.gif" width="16" height="16">Trang chủ,
            	<%End if%>
            	<%if rs("IsCatHomeVote") then%>
            		<img src="../images/update.gif" width="16" height="16">Trang chuyên mục,
            	<%End if%>
            <%end if%>
            Ngôn ngữ: <%=rs("LanguageId")%>
            </em>)</font></td>
        <td align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">
          	<a href="javascript: winpopup('vote_voteItem.asp','<%=rs("VoteId")%>&catid=<%=rs("CategoryId")%>',420,300);"><img src="../images/icon_folder_moderate.gif" border="0" align="absmiddle" title="Thêm/Bớt"></a>
            <a href="javascript: winpopup('vote_edit.asp','<%=rs("VoteId")%>',420,300);"><img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle" title="Sửa"></a>
            <a href="javascript: winpopup('vote_delete.asp','<%=rs("VoteId")%>',300,200);"><img src="../images/icon_closed_topic.gif" width="15" height="15" border="0" align="absmiddle" title="Xóa"></a>&nbsp;&nbsp;</font></td>
    </tr>
    <%j=j+1
          i=i+1
          rs.movenext
       Loop%>
    <tr>
        <td colspan="3" align="right">
            <%Call phantrang(page,rs.pagecount,PAGE_PER_BOOK)
			rs.close
			set rs=nothing
            %>
        </td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
</table>
<%End sub%>

<%Function GetVoteTitle(VoteId)
	Dim rsVote
	set rsVote=server.CreateObject("ADODB.Recordset")
	sql="select VoteTitle from Vote where VoteId=" & VoteId
	rsVote.open sql,con,1
	if rsVote.eof then
		GetVoteTitle=""
	else
		GetVoteTitle=rsVote("VoteTitle")
	end if
	rsVote.close
	set rsVote=nothing
End Function%>

<%Function GetFullDate(theDate)
	ngay = Clng(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = Clng(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = Clng(year(theDate))
	GetFullDate=ngay & "/" & thang & "/" & nam
End Function%>

<%Function GetFullDateTime(theDate)
	gio=Clng(Hour(theDate))
	if gio<10 then
		gio="0" & gio
	end if
	phut=Clng(Minute(theDate))
	if phut<10 then
		phut="0" & phut
	end if
	ngay = Clng(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = Clng(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = Clng(year(theDate))
	GetFullDateTime=gio & ":" & phut & """&nbsp;" & ngay & "/" & thang & "/" & nam
End Function%>

<%Function GetuserEmail(username)
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="select UserEmail from [User] where UserName=N'" & username & "'"
	rs.open sql,con,1
	if not rs.eof then
		GetuserEmail=rs("UserEmail")
	else
		GetuserEmail=""
	end if
	rs.close
	set rs=nothing
End Function%>

<%Function ReplaceHTMLToText(Str)
	Dim strTmp
	strTmp=Trim(Str)
	strTmp=Replace(strTmp,"<br>",chr(13) & chr(10))
	strTmp=Replace(strTmp,"'","''")
	ReplaceHTMLToText=strTmp
End Function%>



<%
' tuannv edit 11/03/2008
Function getChilds(ParentID)
'Tham số đầu vào: ParentCategoryID
'Giá trị trả về: mảng 2 chiều, chứa các thông số: ID, Label, Link của các item tương ứng
'Mục đích: lấy giá trị để tạo submenu

	sql = "SELECT  CategoryName, CategoryID, CategoryLink"
	sql = sql + " FROM NewsCategory"
	sql = sql + " WHERE (LanguageId = 'VN') AND ParentCategoryID = '"&ParentID&"'"
	sql = sql + " ORDER BY CategoryOrder"
	Set rsChild = Server.CreateObject("ADODB.Recordset")
	rsChild.open sql,Con,1
	length = rsChild.recordcount - 1 ' Lấy độ dài của record để gán kích thước cho chiều thứ 2 của mảng LevelOne
	Redim arrChild(2,length) '2=0+1+2 = Id + Label + Link
	i=0
	Do while not rsChild.eof
		arrChild(0,i) = rsChild("CategoryID")
		arrChild(1,i) = rsChild("CategoryName")
		if Trim(rsChild("CategoryLink"))<>"" then 
			arrChild(2,i)=rsChild("CategoryLink")
		else 
			arrChild(2,i)="ShowCat.asp?CatId=" & rsChild("CategoryId")&"&Lang=VN"
		End if
		i = i + 1
		rsChild.movenext
	Loop	
	getChilds = arrChild
	rsChild.close
	set rsChild = nothing
End Function
%>
<% 
function GetCoverPriceNews(NewsID)
	sql		=	"Select Giabia 	from V_News where NewsId="&NewsID
	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		GetCoverPriceNews	=	rsTemp("Giabia")
	else
		GetCoverPriceNews	=	0
	end if
	
	set rsTemp = nothing
end function	
%>

<%
function administrator(f_mesage,user,re_turn)
	administrator = 0
	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	Set rsadm = Server.CreateObject("ADODB.Recordset")
	sqladm	=	"Select * from UserDistribution where UserName='"& user &"'"
	rsadm.open sqladm,con,1				
	txt_mesage	= ""
	if not rsadm.eof then	
		m_editor		= rsadm("m_editor")
		m_order_output	= rsadm("m_order_output")
		m_order_input	= rsadm("m_order_input")
		m_out_store		= rsadm("m_out_store")
		m_user			= rsadm("m_user")
		m_customer		= rsadm("m_customer")
		m_report		= rsadm("m_report")
		m_accounting	= rsadm("m_accounting")
		m_sys			= rsadm("m_sys")
		m_human			= rsadm("m_human")
		m_work			= rsadm("m_work")
		m_cod			= rsadm("m_cod")
		m_sale			= rsadm("m_sale")
		m_ads			= rsadm("m_ads")
		m_faq			= rsadm("m_faq")
		adm				= rsadm("adm")
	rsadm.close
			
	if f_mesage = true then
		select case m_editor
		case 1
			txt_mesage	=	"- Biên tập viên quyền: thêm<br>"
		case 2
			txt_mesage	=	"- Biên tập viên quyền: thêm - sửa<br>"
		case 3
			txt_mesage	=	"- Biên tập viên quyền: thêm - sửa - xóa<br>"
		end select
		select case m_order_output
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý đơn hàng quyền: Xem đơn<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý đơn hàng quyền: Xem - thêm - sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý đơn hàng quyền: Xem - thêm - sửa - xóa<br>"
		end select	
		select case m_order_input
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý nhập hàng quyền: xem<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý nhập hàng quyền: thêm - sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý nhập hàng quyền: thêm - sửa - xóa<br>"
		end select		
		select case m_out_store
		case 1
			txt_mesage	=	txt_mesage+"- Cho phép xuất kho<br>"
		end select		
		select case m_user
		case 1
			txt_mesage	=	txt_mesage+"- Quản trị User quyền: thêm<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản trị User quyền: thêm - sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản trị User quyền: thêm - sửa - xóa<br>"
		end select	
		select case m_human
		case 1
			txt_mesage	=	txt_mesage+"- Quản trị khách hàng: cơ bản<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản trị khách hàng: chi tiết<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản trị khách hàng: quyền gửi email<br>"
		end select	
		select case m_report
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý thông kê: cá nhân<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý thống kê: chuyên sâu<br>"
		end select	
		select case m_accounting
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý kế toán chứng từ: nhập liệu<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý kế toán chứng từ: tổng hợp - kế toán trưởng<br>"
		end select							
		select case m_sys
		case 1
			txt_mesage	=	txt_mesage+"- Ưu tiên quản trị hệ thống<br>"
		end select
		select case m_human
		case 1
			txt_mesage	=	txt_mesage+"- Quản trị nhân sự: cá nhân<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản trị nhân sự: tổng hợp<br>"
		end select
		select case m_work
		case 1
			txt_mesage	=	txt_mesage+"- Quản trị công việc quyền: thêm<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản trị công việc quyền: thêm - sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản trị công việc quyền: thêm - sửa - xóa<br>"
		end select		
		select case m_sale
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý COD quyền: Tiếp nhận phiếu<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý COD quyền: Kiểm soát<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý COD quyền: Thu tiền<br>"
		end select	
		select case m_sale
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý khuyến mãi: thêm<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý khuyến mãi: sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý khuyến mãi: thêm - sửa - xóa<br>"
		end select	
		
		select case m_ads
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý quảng cáo: thêm<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý quảng cáo: sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý quảng cáo: thêm - sửa - xóa<br>"
		end select
					
		select case m_faq
		case 1
			txt_mesage	=	txt_mesage+"- Quản lý hỏi đáp: thêm<br>"
		case 2
			txt_mesage	=	txt_mesage+"- Quản lý hỏi đáp: sửa<br>"
		case 3
			txt_mesage	=	txt_mesage+"- Quản lý hỏi đáp: thêm - sửa - xóa<br>"
		end select	
		select case adm
		case 1
			txt_mesage	=	"- Quyền quản trị cao nhất<br>"
		end select						
		end if	
		Response.Write(txt_mesage)
	end if

	if adm	= 1 then
		administrator	=	5
	else
	select case re_turn
		case "m_editor"
			administrator	=	m_editor
		case "m_order_output"
			administrator	=	m_order_output
		case "m_order_input"
			administrator	=	m_order_input
		case "m_out_store"
			administrator	=	m_out_store			
		case "m_user"
			administrator	=	m_user
		case "m_customer"
			administrator	=	m_customer
		case "m_report"
			administrator	=	m_report
		case "m_accounting"
			administrator	=	m_accounting
		case "m_sys"
			administrator	=	m_sys
		case "m_human"
			administrator	=	m_human
		case "m_work"
			administrator	=	m_work	
		case "m_cod"
			administrator	=	m_cod	
		case "m_sale"
			administrator	=	m_sale	
		case "m_ads"
			administrator	=	m_ads	
		case "m_faq"
			administrator	=	m_faq	
		end select
	end if	
end function

function UserOperation(UserName,txtOperation)
	Day1 = now
	Ngay=Day(Day1)
	Thang=Month(Day1)
	Nam=Year(Day1)
	CreateDate=Thang & "/" & Ngay & "/" & Nam
	CreateDate=FormatDatetime(CreateDate)
	
	txtOldOperation	= ""
	set rsOpera=Server.CreateObject("ADODB.Recordset")
	sql="SELECT * from UserOperation where DATEDIFF(dd,CreateDate,'" & getDateServer() & "')= 0 and UserName='"&UserName&"'"
	rsOpera.open sql,con,1
	if not rsOpera.eof then
		txtOldOperation	=	rsOpera("UserOperation")
		UserOperationID=	rsOpera("ID")
	else
		set rsOperatemp=Server.CreateObject("ADODB.Recordset")
		sql="insert into UserOperation(UserName,UserOperation,CreateDate) values('"&UserName&"',N'"&txtOperation&"','"&getDateServer()&"')"
		rsOperatemp.open sql,con,1
		set rsOperaOpera = nothing
		exit function
	end if
	set rsOpera = nothing
	txtOperation	=	txtOperation&"<br>"&txtOldOperation
	set rsOpera=Server.CreateObject("ADODB.Recordset")
	sql="update UserOperation set UserName='"&UserName&"',UserOperation=N'"&txtOperation&"',LastDate='"&getDateServer()&"' where ID ='"& UserOperationID &"'"
	rsOpera.open sql,con,1	
	set rsOpera = nothing

end function

function OverTimeWork(n_time)
	over_time	= false
	if 	(Weekday(n_time) = 7 and hour(n_time) > 12) or (hour(n_time) >= 18) then
		over_time	= true
	end if
	OverTimeWork	= over_time
end function

function get_ismember(ismember)
	str	=	""
	select case ismember
		case 0
			str =	"Đã nghỉ việc"
		case 1
			str =	"Đang hoạt động"
		case 2
			str =	"Bán thời gian"
		case 3
			str =	"Đối tác"
		case 4
			str =	"Khác"
		case 5
			str =	"Hết hạn hợp đồng"			
	end select
	get_ismember	=	str
end function

function get_staffid(m_user)
	sql="SELECT * from [User] where Username='"& m_user &"'"
	set rsTemp	=	Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,Con,1
	if not rsTemp.eof then
		get_staffid	=	rsTemp("IDNhanVien")
	end if	
end function

sub selectchucdanh(ChucdanhId,pname)
%>
<select name="<%=pname%>" id="<%=pname%>">
    <option value="0">Chọn chức danh</option>
    <%
	sqlChucdanh = "SELECT * FROM Chucdanh"
	Set rsChucdanh = Server.CreateObject("ADODB.Recordset")
	rsChucdanh.open sqlChucdanh,Con,3
	Do while not rsChucdanh.eof
    %>
    <option value="<%=rsChucdanh("ChucdanhID")%>" <%if rsChucdanh("ChucdanhID") = ChucdanhId then%> selected="selected" <%end if%>>
        <%=rsChucdanh("Description")%>
    </option>
    <%
		rsChucdanh.movenext
	Loop
	set rsChucdanh= nothing
    %>
</select>
<%end sub

function getchucdanh(ChucdanhId)
	sqlChucdanh = "SELECT * FROM Chucdanh where ChucdanhID='"&ChucdanhID&"'"
	Set rsChucdanh = Server.CreateObject("ADODB.Recordset")
	rsChucdanh.open sqlChucdanh,Con,3
	if not rsChucdanh.eof then
		getchucdanh	=	rsChucdanh("Description")
	end if
	set rsChucdanh= nothing
end function%>

<%sub selectroom(PhongID,pname)%>
<select name="<%=pname%>" id="<%=pname%>">
    <option value="0">Chọn phòng ban</option>
    <%
	sql = "SELECT * FROM PhongBan"
	Set rsroom = Server.CreateObject("ADODB.Recordset")
	rsroom.open sql,Con,3
	Do while not rsroom.eof
    %>
    <option value="<%=rsroom("PhongID")%>" <%if PhongID=rsroom("PhongID") then %> selected="selected" <%end if%>><%=rsroom("Description")%></option>
    <%
	rsroom.movenext
	Loop
	set rsroom = nothing
    %>
</select>
<%end sub
function getroom(PhongID)
	sql = "SELECT * FROM PhongBan where PhongID='"&PhongID&"'"
	Set rsroom = Server.CreateObject("ADODB.Recordset")
	rsroom.open sql,Con,3
	if not rsroom.eof then
		getroom	=	rsroom("PhongID")
	end if
	set rsroom = nothing
end function%>





<%sub  Fs_Province(key_,idp_)
    sql_  = "SELECT  * FROM [Tb_Provinces] "
    set rsp_ =  Server.CreateObject("ADODB.Recordset")
    rsp_.open sql_,con,1
    IF Not rsp_.EOF THEN 
%>
    <select  name="sclProvince" class="form-control" style="min-width: 200px;max-width: 200px;">
        <option value="-1">-Lựa chọn-</option>
<%
    IF key_  = "Edit" THEN
        do while not rsp_.EOF 
        id_ =  rsp_("id")
        name_ =  rsp_("Name")
        IF idp_ <> "" And idp_ = id_ THEN
            ac = " selected "
        ELSE
            ac= ""
        END  IF
%>
        <option value="<%=id_ %>" <%=ac %> ><%=name_ %></option>
<%                    
        rsp_.MoveNext
        Loop

     ELSE
        do while not rsp_.EOF 
        id_ =  rsp_("id")
        name_ =  rsp_("Name")
%>
        <option value="<%=id_ %>"><%=name_ %></option>
<%
         rsp_.MoveNext
        Loop
     END IF
%>
    </select>
<%
    END IF
    end sub
%>



<%sub  Fs_Hotel(key_,idh_)
    sql1_  = "SELECT  n.Title,n.NewsID FROM  News as n  INNER JOIN NewsDistribution as d on  n.newsID =  d.newsID INNER  JOIN  NewsCategory as c on  d.CategoryID = c.CategoryID"
    sql1_ =  sql1_&" WHERE  c.CategoryLoai = 6 "
    set rsh_ =  Server.CreateObject("ADODB.Recordset")
    rsh_.open sql1_,con,1
    IF Not rsh_.EOF THEN 


%>
    <select name="sclHotel" multiple="multiple" class="form-control" style="min-width: 200px;max-width: 200px;min-height: 207px;">     
        <option value="-1">--Lựa chọn--</option>
<%


    IF key_  = "Edit" THEN
        

        do while not rsh_.EOF 
        'id_ =  rsh_("NewsID")
        id_ =  rsh_("NewsID")
        name_ =  rsh_("Title")

        IF idh_ <> "" And cint(idh_) = cint(id_) THEN
            acc = " selected "
        ELSE
            acc = ""
        END  IF

%>
        <option value="<%=id_ %>" <%=acc %> > - <%=name_ %></option>
<%                    
        rsh_.MoveNext
        Loop

     ELSE
        do while not rsh_.EOF 
        id_ =  rsh_("NewsID")
        name_ =  rsh_("Title")
%>
        <option value="<%=id_ %>"><%=name_ %></option>
<%
         rsh_.MoveNext
        Loop
     END IF
%>
    </select>
<%
    END IF
    end sub
%>




<% 
   sub searchNews(kw)
   end sub  
%>

<%function getDateServer() 

    '9/6/2016 3:37:52 PM-  MM/DD/YYYY H:I:S
    sqltime = "SELECT GETDATE() AS SDateTime"
    Set rsTime=Server.CreateObject("ADODB.Recordset")
	rsTime.open sqltime,con,1
    IF NOT rsTime.EOF THEN 
        getDateServer = Trim(rsTime("SDateTime"))
    ELSE
       ' getDateServer = Now
    END IF
    rsTime.close
	set rsTime=nothing
end function %>




<%Function DelFile(uri_)   
    URPath = Server.MapPath(uri_)  
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    if fs.FileExists(URPath) then
      fs.DeleteFile (URPath)
    end if
    set fs=nothing 
    End Function
%>