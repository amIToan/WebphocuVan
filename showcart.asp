<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="vi" lang="vi-VN">
<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/Constant.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<%
    CateId_ =    Replace(Request.QueryString("cid")," ","+")
    NewsID_ =    Replace(Request.QueryString("NewsID")," ","+")

    CateIdL = Trim(Replace(CateId_,".html",""))
    NewsIDL = Trim(Replace(NewsID_,".html",""))

    IF (CateIdL = "" or IsNull(CateIdL)) AND (NewsIDL = "" or IsNull(NewsIDL)) THEN '  index
        meta_title = page_title
        meta_keyw  = meta_keywords
        meta_desc  = meta_description
    ELSEIF  CateIdL <> "" AND IsNumeric(CateIdL) AND NewsIDL = "" THEN  ' cate
        meta_title =  GetColVal("NewsCategory","CategoryName","CategoryId = '"&CateIdL&"'")  
        meta_keyw  =  GetColVal("NewsCategory","meta_keyword","CategoryId = '"&CateIdL&"'")  
        meta_desc  =  GetColVal("NewsCategory","meta_desc","CategoryId = '"&CateIdL&"'")
	    meta_keyw=meta_keywords&","&meta_keyw
        meta_desc=meta_description&","&meta_desc   
    ELSEIF  CateIdL <> "" AND IsNumeric(CateIdL) AND NewsIDL <> "" AND IsNumeric(NewsIDL)  THEN  ' detail
        meta_title  =   GetColVal("News","Title","NewsId = '"&NewsIDL&"'")    
        meta_keyw   =   GetColVal("V_News","meta_keyword","NewsID = '"&NewsIDL&"'")  
        meta_desc   =   GetColVal("V_News","meta_desc","NewsID = '"&NewsIDL&"'")   
	    meta_keyw   =   meta_keyw&","&meta_keywords
        meta_desc   =   meta_desc&","&meta_description		
    END IF
	
	if len(meta_title)<10 then
        meta_title=meta_title&"-"&page_title
    end if
	
	keyword=Trim(Request.Form("keyword"))
	if keyword ="" then
		keyword	=	Session("keyword")
	end if
	Keyword=Replace(keyword,"'","''")
%>
<head>  
    <title><%=meta_title%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0"/>     
    <link href="/images/logo/icon.ico" rel="icon" type="image/x-icon" />
    <link href="/images/logo/icon.ico" rel="shortcut icon" />
	<meta name="keywords" content="<%=meta_keyw %>"/>
    <meta name="description" content="<%=meta_desc %>"/>
    <link href="/interfaces/liberary/bootstrap4/css/bootstrap.min.css" rel="stylesheet" />
    <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/jquery-3.5.1.min.js"></script>
    <link href="/interfaces/owlcarousel/assets/owl.theme.default.min.css" rel="stylesheet" />
    <link href="/interfaces/owlcarousel/assets/owl.carousel.min.css" rel="stylesheet" />
    <link href="/interfaces/css/sweetalert.css" rel="stylesheet" />
    <link href="/stylesheets/w3style.css" rel="stylesheet" />
    <script type="text/javascript" src="/interfaces/owlcarousel/owl.carousel.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <link href="/interfaces/fonts/font-awesome.css" rel="stylesheet" />
    <link href="/interfaces/css/slick.css" rel="stylesheet" />
    <link href="/interfaces/css/ceo_lam.css" rel="stylesheet" />
    <link href="/interfaces/css/bootstrap.css" rel="stylesheet" />
    <%Call code_google() %>

</head>
<!--#include virtual="/include/func_common.asp"-->
<!--#include virtual="/include/Fs_cotruct.asp"-->
<!--#include virtual="/include/func_tiny.asp"-->
<!--#include virtual="/include/function_toan.asp"-->
<body>
    <%
        Call Header()
        Call Fs_menuMOblie()
        Call Fs_menu()
        Call write_Ads2(IDCate,lang,8,0,0)
        call Orderdetails() 
        IDkey_ = Replace(Request.QueryString("idkey")," ","+")  ' đây chính là tên catorgoryname ứng vs idkey trong wenconfig
        'response.write(IDkey_&"<br>")
        IDCate = Replace(Request.QueryString("cid")," ","+") ' đây chính là tên catorgoryId ứng vs cid trong wenconfig
        'response.write(IDCate&"<br>")
        IDNews = Replace(Request.QueryString("newsid")," ","+")
        'response.write(IDNews&"<br>"&"lolloolol")
            IF IDCate <> "" AND  InStr(IDCate,".html") > 0  THEN
                IF IsNumeric(Trim(Replace(IDCate,".html",""))) THEN 'Is IsNumeric		
                IDC = Trim(Replace(IDCate,".html",""))' đây chính là catergoryID 
                'Response.write(IDC) 
                cLoai_= getColVal("NewsCategory","CategoryLoai","CategoryID = '"&IDC&"'")
                'Response.write(cLoai_)                              
                  SELECT CASE cLoai_
                       CASE "1"
							   Call Fs_CateNews(IDC,Lang)         
                       CASE "2"
                                call commonProducts(IDC,cLoai_,Lang)   
                       CASE "3"
                                Call Fs_CateNews(IDC,Lang)
                       CASE "4"                           
                                Call Fs_CateNews(IDC,Lang)
                       CASE "5"
                                Call Fs_CateNews(IDC,Lang)
                        CASE "7"
                                Call Fs_CateNews(IDC,Lang)
                       CASE "8"
                                Call Fs_Paner(799,Lang)  
                       CASE "9"
                                Call Fs_Faq()
                               
                       CASE "10"
                               Call Fs_CateService(IDC,2) 
                               Call Fs_ServiceCommon(2,2)
                       CASE ELSE   
                        Response.Redirect("/")                           
                   END SELECT
                ELSE
                   'search
				    keyword=Trim(Request.Form("keyword"))
					Keyword=Replace(keyword,"'","''")
                    keywordMobile=Trim(Request.Form("keywordMobile"))
					keywordMobile=Replace(keywordMobile,"'","''")
                    if Keyword="" then 
                        keyword = keywordMobile
                    end if 
                    Call SearchByTitle(Keyword)
                END IF
            ELSE
                Response.Redirect "/"
            END IF
       Response.write ("</div>")
    %>
    <%Call  Fs_Footer() %>
    <%Call backTop() %>
    <%  Call Item_support_Toan(lang) %>
    <script type="text/javascript" src="/interfaces/js/custom.js"></script>
    <script type="text/javascript" src="/interfaces/js/ajax-asp.js"></script>
    <script type="text/javascript" src="/interfaces/js/sweetalert.min.js"></script>
	<script type="text/javascript" src="/interfaces/owl.carousel/owl.carousel.min-2.3.4.js"></script>

    <!-- Messenger Chat plugin Code -->
    <div id="fb-root"></div>

    <!-- Your Chat plugin code -->
    <div id="fb-customer-chat" class="fb-customerchat">
    </div>

    <script>
      var chatbox = document.getElementById('fb-customer-chat');
      chatbox.setAttribute("page_id", "507888815947986");
      chatbox.setAttribute("attribution", "biz_inbox");
    </script>

    <!-- Your SDK code -->
    <script>
      window.fbAsyncInit = function() {
        FB.init({
          xfbml            : true,
          version          : 'v12.0'
        });
      };

      (function(d, s, id) {
        var js, fjs = d.getElementsByTagName(s)[0];
        if (d.getElementById(id)) return;
        js = d.createElement(s); js.id = id;
        js.src = 'https://connect.facebook.net/vi_VN/sdk/xfbml.customerchat.js';
        fjs.parentNode.insertBefore(js, fjs);
      }(document, 'script', 'facebook-jssdk'));
    </script>
    <script src="/javascript/cus_option.js"></script>
</body>
</html>


            
            