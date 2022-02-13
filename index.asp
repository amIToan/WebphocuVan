<html>
<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/Constant.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<%
    CateId_ =   Replace(Request.QueryString("cateId")," ","+")
    NewsID_ =   Replace(Request.QueryString("NewsID")," ","+")
    CateIdL =   Trim(Replace(Request.QueryString("cateId")," ","+"))
    NewsIDL =   Trim(Replace(NewsID_,".html",""))
	
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
        meta_title      =  GetColVal("News","Title","NewsId = '"&NewsIDL&"'")    
        meta_keyw  = GetColVal("V_News","meta_keyword","NewsID = '"&NewsIDL&"'")  
        meta_desc  = GetColVal("V_News","meta_desc","NewsID = '"&NewsIDL&"'")   
	    meta_keyw=meta_keyw&","&meta_keywords
        meta_desc=meta_desc&","&meta_description	
    END IF
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
    <%=embed_head %>
    <%call code_google() %>
</head>
<!--#include virtual="/include/func_common.asp"-->
<!--#include virtual="/include/Fs_cotruct.asp"-->
<!--#include virtual="/include/func_tiny.asp"-->
<!--#inclue virtual="/include/function_toan.asp" -->
<body>
<div id="fb-root"></div>
<script async defer crossorigin="anonymous" src="https://connect.facebook.net/en_GB/sdk.js#xfbml=1&version=v12.0" nonce="vJa849RD"></script>
    <%
        IDCate = Replace(Request.QueryString("cateId")," ","+")
        IDNews = Replace(Request.QueryString("newsid")," ","+")
        IDN = Trim(Replace(IDNews,".html",""))
        IF IDCate  <> ""  And IsNumeric(Trim(Replace(IDN,".html","")))   THEN'view-detail
	    Call UpdateNewsCounter(IDN)
        Call Header()  
        Call Fs_menuMOblie()
        Call Fs_menu() 
        'Call write_Ads2(IDCate,lang,8,0,0)
        call Orderdetails()
            cLoai_= getColVal("NewsCategory","CategoryLoai","CategoryID = '"&IDCate&"'")
            IF cLoai_  <> "" THEN
                SELECT CASE cLoai_
                    CASE "1"
                        Call Fs_NewsDetail_hoangduc_makeup(IDN)
                        Call Fs_NewsInvolve(IDCate,IDN)
                    CASE "2"
                        Call Fs_NewsDetail_Product(IDN,cLoai_)
                    CASE "3"
                        Call Fs_NewsDetail_hoangduc_makeup(IDN)
                        Call Fs_NewsInvolve(IDCate,IDN)              
                    CASE "4"
                        Call Fs_NewsDetail_hoangduc_makeup(IDN)
                        Call Fs_NewsInvolve(IDCate,IDN)
                    Case "5"
                        Call Fs_NewsDetail_hoangduc_makeup(IDN)
                        Call Fs_NewsInvolve(IDCate,IDN)
                    CASE "6"
                            Response.Write  "Đang cập nhật."
                    Case "7"
                        Call Fs_NewsDetail_hoangduc_makeup(IDN)
                        Call Fs_NewsInvolve(IDCate,IDN) 
                    CASE "9"
                               Call Fs_Faq()        
                    CASE ELSE   
                        Call Fs_information(IDN)                    
                END SELECT
            ELSE 
               'Response.Redirect("/")
            END IF
        ELSE'view-index
		 'Response.Write  "Đang cập nhật."
         %>
         <div style="background-image: url('/images/logo/background.png'); display: flow-root">
         <%
          Call Header() 
          Call Fs_menuMOblie()
          Call Fs_menu()
          %>
          </div>
          <%
          Call write_Ads2(IDCate,lang,8,0,0)
          %>
           <div style="background-image: url('/images/logo/background.png'); display: flow-root">
          <%
          Call NewsHome()
          'call NewsandLibary()
          'call Getfeedback()
          call Orderdetails()                
        END IF  
		%>
        </div>
        <div>
        <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3723.2464554632356!2d105.79770881534779!3d21.062816585980002!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x3135ab08310865fd%3A0xd60aa33e9ae60d33!2zVGhlIE5vb2RsZSBIb3VzZSAtIFBo4bufIGPhu6dhIG5nxrDhu51pIFRyw6BuZyBBbg!5e0!3m2!1svi!2s!4v1644766846590!5m2!1svi!2s" width="100%" height="450" style="border:0;" allowfullscreen="" loading="lazy"></iframe>
        </div>
        <%      
        Call  Fs_Footer()
       %>
    <%call backTop() %>
    <%  Call Item_support_Toan(lang) %>
    <!-- Goi ra moblie bars -->
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
    <script src="/javascript/shopping-cart.js"></script>
    <script src="/javascript/cus_option.js"></script>
</body>
</html>

