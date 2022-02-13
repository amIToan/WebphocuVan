<!--#include virtual="/include/config.asp" -->
<%
    lang = Session("Language")
    IF lang = ""  or lang = "VN"  THEN 
      '  lang = "VN"
      '  IF lang = "VN" THEN

            Faq_lang  = "VN"
            view_detail = "xem chi tiết..." 
            v_detail = "xem chi tiết..." 
            isHome  = "Trang chủ"
            msg_faq = "Gửi câu hỏi - Thắc mắc của bạn hãy để tuduyketoan giải đáp giúp bạn. "
            st_send = "Gửi"
            st_cancel = "Hủy bỏ"
            ms_contact= "Liên hệ với chúng tôi:"


            ms_hoten = "Họ và tên:"
            ms_email = "Địa chỉ Email:"
            ms_title = "Tiêu đề:"
            ms_content = "Nội dung câu hỏi:"


            ms_tk = "Thống kê"
            ms_isonline = "Đang online:"
            ms_today    = "Hôm nay:"
            ms_tongac    = "Tổng truy cập:"

            ms_search =  "Kết quả tìm kiếm với từ khóa:"

            ms_info="THÔNG TIN CÔNG TY"
    

        ELSE
            Faq_lang  = "EN"
        
            view_detail = "Move next...." 
            v_detail = "Move next...." 
            isHome  = "Home"
            msg_faq = "Send a question - your question let tuduyketoan answers help you."
            st_send = "Send"
            st_cancel = "Cancel"
            ms_contact= "contact us : "

            ms_tk = "Statistics"
            ms_isonline  = "Is online: "
            ms_today     = "Today: "
            ms_tongac    = "Total: "

            ms_search =  "Keyword search:"

        END IF
    'END IF



    if lang = "" then lang = "VN"
    Set rs=Server.CreateObject("ADODB.Recordset")		
	sql="select top 1 * from Company where lang = '"&lang&"'"
  '  Response.Write sql
	rs.open sql,con,1
	if not rs.eof then
		company		        =	Trim(rs("company"))
		Tel_sys			    =	Trim(rs("Tel"))
		Hotline			    =	Trim(rs("Hotline"))
		Bank			        =	Trim(rs("Fax"))
		Email		        =	Trim(rs("Email"))
		Website		        =	Trim(rs("Website"))
		Address	            =	Trim(rs("address"))
		calltime	        =	Trim(rs("calltime"))
		STK	         		=	trim(rs("Masothue"))
		Tentaikhoan		    =	trim(rs("GPKD"))
		page_title	        =	trim(rs("page_title"))
		TitleF	        	=	trim(rs("TitleF"))
		meta_description	=	trim(rs("meta_description"))
		meta_keywords		=	trim(rs("meta_keywords"))
		icon				=	trim(rs("icon"))
		Logo				=	trim(rs("Logo"))
		logoF				=	trim(rs("logoF"))
		background			=	trim(rs("background"))
		banner1				=	trim(rs("banner"))
		footer1				=	trim(rs("footer"))		
		Cfont				=	trim(rs("Cfont"))
		fsize				=	rs("fsize")
		idgoogle			=	rs("idgoogle")
        idfacebook          =   rs("idfacebook")
        idgplus             =   rs("idgplus")
        idskype             =   rs("idskype")
        idEmail             =   rs("idEmail")
        idvideo             =   rs("idvideo")
        idyoutube           =   rs("idyoutube")

		introduction		=	Trim(rs("introduction"))
		show_intro_home		=	rs("show_intro_home")

        embed_head		    =	rs("embed_head")
        embed_footer		=	rs("embed_footer")

   ' Response.Write Logo
	end if
    sql = ""
	rs.close
	set rs=nothing
    Cfont="arial"
	defaultfont = 14+fsize	

    

	VarPort=""
	NewsImagePath = "/images_upload/"
	DirectoryStored="/images_upload/"
	AudioVideoPath="/images_upload/"	
	Path=server.MapPath(NewsImagePath)

	DirectoryStored="/Document_upload/"
	PathDoc=server.MapPath(DirectoryStored)

Function PhanQuyen(Quyen)
	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
End Function

Function GetNumeric(sValue,DefaultValue)
	Dim intValue
	if not IsNumeric(sValue) or trim(sValue)="" then
		intValue=DefaultValue
	else
		intValue=Clng(sValue)
	end if
	GetNumeric=intValue
End Function%>

<%Sub code_google() %>
<% if idgoogle<>"" then %>
	<!-- Global site tag (gtag.js) - Google Analytics -->
	<script async src="https://www.googletagmanager.com/gtag/js?id=<%=idgoogle%>"></script>
	<script>
		window.dataLayer = window.dataLayer || [];
			function gtag(){dataLayer.push(arguments);}
				gtag('js', new Date());
		gtag('config', '<%=idgoogle%>');
	</script>
<%end if %>
<%End Sub %>


