<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
    <link href="/interfaces/css/bootstrap.css" rel="stylesheet" />
<%
Sub Fs_NewsDetail_hoangduc_makeup(NewsID) 
    sqlN = "SELECT * FROM V_News WHERE   status = '4' And  NewsID = '"&NewsID&"' And LanguageID='"&lang&"' order by LastEditedDate DESC"
    set rsN = Server.CreateObject("ADODB.Recordset")
    rsN.open sqlN,con,1
    If not rsN.eof then
        Title       = Trim(rsN("Title"))
        desc        = Trim(rsN("Author"))
        FTitle      = Trim(rsN("Description")) 
        Ncontent    = Trim(rsN("Body")) 
        cateID      = Trim(rsN("CategoryID"))
        NewsID      = Trim(rsN("NewsID"))    
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&cateID&"'")
        urlvideo = Trim(rsN("url_video"))
         f_bd = "frameborder='0'  allowfullscreen "
        str_ = InStr(1,urlvideo,"https://www.youtube.com/watch?v=") 'xác định link youtube. Nó start từ 1 và nó tìm xem chuỗi đó có nó chứa hay không và vị trí bắt đầu chuỗi str2
            IF urlvideo <> "" And   str_ > 0 THEN 
                idvd_ = Trim(Replace(urlvideo,"https://www.youtube.com/watch?v="," "))
          END IF
          w_ = "100%"
%>
<div class="container-fluid product_content">
    <div class="container">
        <%=GetListParentCatNameOfCatId2(cateID,NewsID)%>
      <div class="col-xs-12 col-sm-12 col-md-7 col-lg-9 ">
        <h1 class="intro-title text-center"><%=Title %></h1>
        <div class="intro-body" style="font-size: 12pt"><%=Ncontent %></div>
        <% if idvd_ <> "" then%>
        <div class="d-flex justify-content-center pb-3">
                <iframe width="<%=w_ %>" style="aspect-ratio:4/3"  src="https://www.youtube.com/embed/<%=idvd_ %>" <%=f_bd %>></iframe>
        </div>
        <%End if%>
      </div>
      <div class="col-xs-12 col-sm-12 col-md-7 col-lg-3 w3-padding-left">
        <% call siderbar(15,cateID) %>
      </div>
    </div><!--/.container--->
</div>
<%
    END IF 'not eof
    rsN.Close
%>
<%End Sub %>
<% Sub search(key)%>
<%
  keyword=Trim(Request.Form("keyword"))
  ' bắt name của form về qua phương thức post
  if keyword ="" then
    keyword = Session("keyword")
  end if

  Keyword=Replace(keyword,"'","''")

  seach_filter=Request.Form("select_filter")
  if seach_filter="" then
    seach_filter=Session("seach_filter")
  end if

%>
<div class="container">
<div class="w3-row">
  <div class="form-group" style="position: relative; height: 26px;">
        <div style="float: left; background: #75bb1f;" class="in-title">
            <h3 class="H-Title">Tìm kiếm</h3>
        </div>
  </div>
  <hr class="Hr-Title">

    <h4><strong style="color: #e53600;">Tìm kiếm với từ khóa:  <span class="icon-key"></span><%=Keyword %></strong></h4>
</div>
 <div id="search_hoangduc" class="col-xs-12 col-sm-12 col-md-7 col-lg-9 w3-padding-right">
<script type="text/javascript">
    function ShowHelpSearch()
    {
  
        if(document.fNews.hShowHelp.value==0)
        {
            document.getElementById("ShowHelp").style.display="";
            document.fNews.hShowHelp.value=1;
        }
        else
        {
            document.getElementById("ShowHelp").style.display="none";
            document.fNews.hShowHelp.value=0  ;
        }
    }
    function ONBorder()
    {
        document.getElementById("keywordDetail").style.border='#99CCFF solid 3';
    }
    function OFFBorder()
    {
        document.getElementById("keywordDetail").style.border='#99CCFF solid 1';
    }
</script>
<%if Keyword<>"" then%>
            <%     
  Keyword=replace(Trim(Keyword),"  "," ") 
  StringSearch=replace(Trim(Keyword),"   "," ")
  StringSearch=replace(Trim(StringSearch),"  "," ")
  StringStaus = StringSearch
  if seach_filter <> 0 then
    select case seach_filter
    case 1
      Loai =  "Tìm theo tiêu chí"   
    case 2
      loai  = "Tiêu đề"
    case 3
      loai  = "Nội dung"
    case 4
      loai  = "Nhà xuất bản"
    case 5
      loai  = "Tác giả"
    case 6
      loai  = "Giá gốc"
    case 7
      loai  = "Tất cả"
    case 8
      loai  = "Mã sản phẩm"
    case 9
      loai  = "Loại sản phẩm"
    end select
  end if  
  Session("seach_filter")=seach_filter
    Session("keyword")=Keyword
  page  = GetNumeric(Request.QueryString("page"),1) 
  if page > 1 then
  ' nếu trang lớn hơn 1 có nghĩa là có phân trang thì

    set dictSearch  = Session("dictSearch")
    ' gán biến dictSearch bằng Session("dictSearch") bắt về

  else

  Set dictSearch  =   Server.CreateObject("Scripting.Dictionary") 
  ' kết nối đến từ điển Scripting.Dictionary

  Set rs = server.CreateObject("ADODB.Recordset")
  ' gán biến kết nối đến DB

  sql = "SELECT NewsID " &_
    "FROM V_News "

  select case seach_filter
    case 1
      StringTitle = UCASE(StringSearch)

      sql = sql+ "WHERE (Title like N'%"&StringTitle&"%' ) order by CreationDate desc" 
      'response.write(sql)   
    case 7
      StringTitle = UCASE(StringSearch)
      sql= sql+ "WHERE ({fn UCASE(Title)} like N'%"&StringTitle&"%' OR ({fn UCASE(CategoryName)} like N'%"&StringTitle&"%') or Description like N'%"&StringTitle&"%' OR Body like N'%"&StringTitle&"%')"
    case 2
      StrTitle  = UCASE(StringSearch)
      sql= sql+ " Where ({fn UCASE(Title)} like N'%"&StrTitle&"%')"
    case 3
      sql= sql+ "Where (Body like N'%"&StringSearch&"%')"
    case 8
      sql= sql+ "Where (idcode like N'%"&StringSearch&"%')"
    case 9
      StringTitle = UCASE(StringSearch)       
      sql= sql+ "Where ({fn UCASE(CategoryName)} like N'%"&StringSearch&"%')"
  end select
    ' sql= sql + "ORDER BY NewsID DESC"
  rs.open sql,con,3
  
  do while not rs.eof
    ' nếu có bản ghi
    iItemID = rs("NewsID")
    if dictSearch.Exists(iItemID)=false then
      dictSearch.Add iItemID, StringSearch
    end if
    if dictSearch.count > 501 then
      exit do
    end if                
    rs.movenext
  loop  
 
  set rs=nothing
  end if  
  
  if dictSearch.count = 0  then
    '   Khong dấu
    StringSearch  = UCase(MID(Keyword,1,2))
    Set rs=server.CreateObject("ADODB.Recordset")
    sql="SELECT * FROM V_News "
    sql= sql+ "WHERE ({fn UCASE(Title)} like N'%"&StringSearch&"%')"
    sql= sql+ " order by NewsID Desc" 
    rs.open sql,con,3 
    Response.Write(rs.recordcount) 
    Do while not rs.eof
      select case seach_filter
        case 1,2
          Title = rs("Title")
        case 4
          Title = rs("nxb")
        case 5
          Title = rs("tacgia")
      end select
      if len(Title) >= len(Keyword) then
        iPos  = InStr(Uni2NONE(UCASE(Title)),Uni2NONE(UCASE(Keyword)))
        ' UCASE() biến đổi chữ thành in hoa
      iF iPos >= 1 then
        iItemID = rs("NewsID")
        StringSearch  = Mid(rs("Title"),iPos,Len(Keyword))
        if dictSearch.Exists(iItemID)=false then
          dictSearch.Add iItemID, StringSearch
        end if          
      end if
      end if    
      rs.movenext
    loop
    set rs=nothing
  end if
  if seach_filter=1 or seach_filter=2 then    
    if dictSearch.count = 0  and len(StringSearch) <= 30 then
      ' Tương tự và không giấu
      StringSearch  = UCase(MID(Keyword,1,2))
      Set rs=server.CreateObject("ADODB.Recordset")
      sql="SELECT NewsID,Title FROM V_News "
      sql= sql+ "WHERE ({fn UCASE(Title)} like N'%"&StringSearch&"%') order by NewsID Desc" 
      rs.open sql,con,3 
      Set SortSearch  =   Server.CreateObject("Scripting.Dictionary")
      Do while not rs.eof and len(Keyword) <= 30
        if len(rs("Title")) >= len(Keyword) then
          wordPercent = simalarString(UCASE(rs("Title")),UCASE(Keyword))
        iF wordPercent >= 90 then
            iItemID = rs("NewsID")
            StringSearch  = rs("Title")
            if dictSearch.Exists(iItemID)=false then
              dictSearch.Add iItemID, StringSearch
              SortSearch.add iItemID, wordPercent
            end if          
          end if
        end if    
        rs.movenext
      loop
    if dictSearch.count = 0  then
      on error Resume Next
      rs.MoveFirst
      Do while not rs.eof and len(Keyword) <= 30
        if len(rs("Title")) >= len(Keyword) then
          wordPercent = simalarString(Uni2NONE(UCASE(rs("Title"))),Uni2NONE(UCASE(Keyword)))
          iF wordPercent >= 80 then
            iItemID = rs("NewsID")
            StringSearch  = rs("Title")
            if dictSearch.Exists(iItemID)=false then
              dictSearch.Add iItemID, StringSearch
              SortSearch.add iItemID, wordPercent
            end if          
          end if
        end if    
        rs.movenext
      loop
    end if
      
    SortDictionary SortSearch,2,1 
    Set dictSortTemp  =   Server.CreateObject("Scripting.Dictionary")   
    For Each Key in SortSearch
      dictSortTemp.add Key,dictSearch(Key)
    next
    dictSearch.RemoveAll
    set dictSearch  = dictSortTemp      
    set rs=nothing
    end if
    if  dictSearch.count <= 0 then
      ' từ có ý nghĩa và không dấu
      arKeyword = split(TRim(Keyword)," ")
      for iKey1 = 0 to Ubound(arKeyword)
        for iKey2 = iKey1 + 1 to Ubound(arKeyword)
          StrTitle  = UCASE(arKeyword(iKey1) + " " + arKeyword(iKey2))
          StringSearch  = UCase(MID(StrTitle,1,2))
          Set rs=server.CreateObject("ADODB.Recordset")
          sql="SELECT NewsID,Title FROM V_News "
          sql= sql+ "WHERE ({fn UCASE(Title)} like N'%"&StringSearch&"%') order by NewsID Desc" 
          rs.open sql,con,3   
          Do while not rs.eof
            iPos  = InStr(Uni2NONE(UCASE(rs("Title"))),Uni2NONE(UCASE(StrTitle)))
            iF iPos >= 1 then
              iItemID = rs("NewsID")
              StringSearch  = Mid(rs("Title"),iPos,Len(StrTitle))           
              if dictSearch.Exists(iItemID)=false then
                dictSearch.Add iItemID, StringSearch
              end if          
            end if
            if dictSearch.count > 501 then
              exit do
            end if          
            rs.movenext
          loop
          set rs=nothing            
        next
      next        
    end if
  end if
  stemp = StringStaus&"("&dictSearch.count&")"
  sql1="insert into StatusSearch(Keyword,Category) values(N'"& stemp &"',N'"& loai &"')"
  Set rs1 = Server.CreateObject("ADODB.Recordset")
' rs1.open sql1,Con,3
  set rs1 = nothing
  
  Call ReSearchOther(dictSearch,page)
            %>
            <%if Round(dictSearch.count/20) > 0 then
            Response.Write("<div class='w3-row'>")
                Response.Write("<div class='w3-bar  w3-margin-top w3-margin-bottom'>")
                    Response.Write("<strong><u>Trang:</u></strong>") 
                        'Xu ly' URL QUERY_STRING
                        If Request.ServerVariables("SERVER_PORT")=80  Then
                            strUrlRoot = "http://" & Request.ServerVariables("HTTP_HOST")
                        Else
                            strUrlRoot = "https://" & Request.ServerVariables("HTTP_HOST")
                        End If

                        url_link       =  "/"&keyword&".html"

                        sURL=strUrlRoot&url_link&"/page-" 
        Set Session("dictSearch") = dictSearch 
        For i=1 to Round(dictSearch.count/20)
                    if i<>clng(page) then     
                        if i=1 then
                    response.Write "<a class=""w3-button"" href=""" & strUrlRoot&url_link & """>" & i & "</a>"                
                        else               
                    response.Write "<a class=""w3-button"" href=""" & sURL & i & """>" & i & "</a>"
                        end if
                else      
                        Response.write "<a class=""w3-button w3-green"">"&i&"</a>"              
                end if               
              Next     
                Response.Write("</div><!---/.w3-bar---></div>")   
                ' response.write("</ul>") 
            end if%>
            <!-- </ul>  -->
      </div> 
      <div class="col-xs-12 col-sm-12 col-md-5 col-lg-3">
        <%
          call siderbar(10)
        %>
    </div>    
  </div>


<%end if%>
<%End Sub%>

<% Function ReSearchOther(dictSearch,page)%>
  <ul class="w3-ul w3-hoverable w3-white w3-row">     
<%  
  iCol=1
  ' gán biến iCol = 1

  kt = 20*(page-1)+1

  stt = 1
  ' đếm 
  
  For Each Key in  dictSearch
    ' gán biến tất cả giá trị trả về của dictSearch => Key

  if stt>=kt then
    if kt > page*20 then
      exit for
    end if
    Set rs=server.CreateObject("ADODB.Recordset")
    sql="SELECT * " &_
    "FROM V_News where NewsID="&Key&" AND LanguageID='"&Lang&"'"
    rs.open sql,con,3
    if not rs.eof then
      Description=rs("Description")
      Title =rs("Title")
      PictureAlign  = rs("PictureAlign")
      if PictureAlign = "" then
        PictureAlign = "left"
      end if
        ' if len(Title) > 30 then
          ' Title1 = left(Title,120)&"..."
      Description=Replace(Description,LCase(dictSearch(Key)),"<font style=""BACKGROUND-COLOR: #FFFF66"">" & LCASE(dictSearch(Key)) & "</font>")

      ' if Title1  = lcase(Title)then

      Title3 = Replace(Title,uCase(dictSearch(Key)),"<font style=""BACKGROUND-COLOR: #FFFF66"">" & UCase(dictSearch(Key)) & "</font>")
        ' else if Title1  = ucase(Title)then
       ' Title2= Replace(Title1,uCase(dictSearch(Key)),"<font style=""BACKGROUND-COLOR: #FFFF66"">" & uCase(dictSearch(Key)) & "</font>")

     ' end if
        ' end if 
     ' end if
        %>
        <li class="w3-padding-8 w3-half w3-border">
            <a href="/<%=rs("NewsId")%>/<%=rs("CategoryId")%>/<%=Replace(Uni2NONE(Title)," ","-")+".html"%>">
              <%Call ShowPicture(rs("PictureId"),PictureAlign,10,rs("NewsId"),rs("CategoryId"),Replace(Uni2NONE(rs("title"))," ","-"))%>
               
            </a>
            <p class="" style="margin:0;height: 100px;text-align:left;overflow:hidden">
                <a class="sl-news-Tl" href="/<%=Replace(Uni2NONE(Title)," ","-")%>/<%=rs("CategoryId")%>/<%=rs("NewsId")%>.html">
                  <%=Title3%>
                </a>

            </p>
            <a href="/<%=rs("NewsId")%>/<%=rs("CategoryId")%>/<%=Replace(Uni2NONE(Title)," ","-")+".html"%>" class="w3-text-red w3-right">
                <% '=getLang(Lang,162)%>
            </a>
        </li>       
        <%
      iCol = iCol + 1

    end if
    kt = kt +1

  end if
  stt = stt+1
  next      
        %>
<!-- </ul> -->
<%  
End Function %>

<%
Sub UpdateNewsCounter(NewsID)
    'Kiểm tra Cookies để chống refresh
    Cookies_NewsId=GetNumeric(Request.Cookies("XSEO"),0)
    if Cookies_NewsId <> NewsId  then
        'Nếu hợp lệ Gọi UpdateNewsCounter
        Dim cm
        set cm = CreateObject("ADODB.Command")
        cm.ActiveConnection = strConnString

        cm.commandtype=4 'adstoredProc
        cm.CommandText = "UpdateNewsCounter"
        cm.Parameters.Append cm.CreateParameter("IPAddress", 200, 1, 15, Request.ServerVariables("REMOTE_ADDR"))
        cm.Parameters.Append cm.CreateParameter("NewsID", 3, 1, 8, NewsID)
        'Set objparameter=objcommand.CreateParameter (name,type,direction,size,value)
        'cm.Parameters.Append cm.CreateParameter("counter", 3, 2, 4)
        cm.execute()
        set cm=nothing
        'Ghi lại Cookies
        Response.Cookies("XSEO")=NewsId
        Response.Cookies("XSEO").Expires=DateAdd("n",2,now()) 'Chỉ lưu Cookies trong 2 phút
    end if
    
End Sub%> 

<%
    sub Header()
   
%>
<!--<div class="container-fluid bkg-dark" >
    <div class=" header_top container w3-padding">
        <div class="header-content-log"> <%if Hotline <> "" then%> <a class="w3-text-white" href="<%=Hotline %>">Hotline: <% =Hotline %></a> <%end if%></div>
        <div class="header_top_middle w3-hide-small w3-text-white "><span class ="header_top_middle_animated">Chào mừng bạn đến với Phở Tràng An!</span></div>
        <div class="header_top_button w3-hide-small">
                <button class="w3-button w3-green w3-round-large w3-text-white" disabled>Đăng nhập</button>
                <button class="w3-button w3-green w3-round-large w3-text-white" disabled>Đăng ký</button>
        </div>
    </div>
</div>-->
<div class="header-content-container" >
    <div class="header-content container w3-section">
                <div class="header-content-log">
                    <a href="/" target="_blank" >
                        <img src="/images/logo/<%=logo%>" alt="<%=company%>" title="<%=company%>" style="max-width:150px;" />
                    </a>
                </div>
                <div class="header-content-files w3-hide-small">
                    <a href="https://file.hstatic.net/1000382335/file/0710_-_tltm_lauching-npp-tiengvietol_02_57f99448ae0e49ee96447b49aee22572.pdf" target="_blank" >
                        <img src="/administrator/images/down.jpg" alt="download">
                    </a>
                    <p><b>Download</b><p>
                    <p>Bộ tài liệu giới thiệu<p>
                </div>
                <div class="header-content-time w3-hide-small">
                    <p><b>Thời gian làm việc</b><p>
                    <p>T2 - T7 : Giờ hành chính<p>
                </div>
                <div class="header-content-buying">
                    <form class="d-flex">
                            <button class="btn1 btn-outline-dark shoppingCartButton" type="button">
                                <i class="fas fa-cart-plus me-1"></i>
                                Cart
                                <span class="badge bg-dark text-white ms-1 rounded-pill" id="sum-prices">0</span>
                            </button>
                    </form>
                </div>
    </div>
    <form action="/timkiem.html" method="post" name="fTimKiemNhanh" id="fTimKiemNhanh">
        <div class="search-box w3-hide-small">
            <input type="text"  placeholder="Type to search" class="search-txt" id="keyword" name="keyword">
                        <button class="search-btn">
                            <i class="fas fa-search"></i>
                        </button>
        </div>
    </form> 
</div>
<%        
    end sub
%>
    <!-- Viet lai -->
      <%
    sub Fs_menu()
%>
<nav class="container-fluid hide-medium w3-hide-small ">
        <ul class="container self-list ">
              <% dem =0  
                         sqlMenu="SELECT    * " &_
                             "FROM      NewsCategory " &_
                             "WHERE   LanguageId = '"&lang&"' and  (ParentCategoryID = 0) and (CategoryStatus = 2)   AND NOT CategoryHome='8' ORDER BY CategoryOrder"

                     set rsMenu = Server.CreateObject("ADODB.RECORDSET")
                     rsMenu.Open sqlMenu,con,1
               	do while not rsMenu.eof
                    CategoryID  =   rsMenu("CategoryId") 'no sẽ lấy hết tất các catergory ID
                   
                    if Trim(rsMenu("CategoryLink"))<>"" then 'Nó sẽ trim nững khoảng chống của Link 
                        varLink=rsMenu("CategoryLink") ' ví dụ ở đây nó sẽ là varLInk = /index.asp
                    else 

                        varLink= "/"&Replace(Uni2NONE(rsMenu("CategoryName"))," ","")&"/"&rsMenu("CategoryId")&".html" 'khi mà nó bằng rỗng thì sẽ là 
                  '/Gioithieu/807.html

                    End if 
                    query = trim(Replace(Request.QueryString("cateId")," ","+"))
                    if query <> "" then
                            if trim(rsMenu("CategoryId")) = query then active = "active-link" else active = ""
                    else
                            if dem = 0 then active = "active-link" else active = "" 
                                dem = dem +1
                    end if
					YoungestChildren = rsMenu("YoungestChildren")
					CategoryID = rsMenu("CategoryID")
					CategoryName = rsMenu("CategoryName")                                                
					call subMenu(YoungestChildren,CategoryID,varLink,CategoryName,active)     

				rsMenu.MoveNext()
				loop
%>
            </ul>         
</nav>
<%
    end sub
%>
      <!-- Moblie bars -->
        <%
    sub Fs_menuMOblie()


     CID = trim(Replace(Request.QueryString("cid")," ","+"))

    IF CID <> "" THEN 
        acindex = ""
    ELSE
        acindex = "active"        
    END  IF
%>
      <!-- moblie bars -->
            <nav class="w3-hide-large w3-hide-medium" >
                <div class="burger"><i class="fas fa-bars"></i></div>
                <div class="nav__menu show-burger">
                    <ul class="nav__list ">
                         <% dem =0  
                         sqlMenu="SELECT    * " &_
                             "FROM      NewsCategory " &_
                             "WHERE   LanguageId = '"&lang&"' and  (ParentCategoryID = 0) and (CategoryStatus = 2)   AND NOT CategoryHome='8' ORDER BY CategoryOrder"

                     set rsMenu = Server.CreateObject("ADODB.RECORDSET")
                     rsMenu.Open sqlMenu,con,1
               	do while not rsMenu.eof
                    CategoryID  =   rsMenu("CategoryId")
                   
                    if Trim(rsMenu("CategoryLink"))<>"" then 
                        varLink=rsMenu("CategoryLink")
                    else 

                        varLink= "/"&Replace(Uni2NONE(rsMenu("CategoryName"))," ","")&"/"&rsMenu("CategoryId")&".html"

                    End if 
                    query = trim(Replace(Request.QueryString("cateId")," ","+"))
                    if query <> "" then
                            if trim(rsMenu("CategoryId")) = query then active = "active-link" else active = ""
                    else
                            if dem = 0 then active = "active-link" else active = "" 
                                dem = dem +1
                    end if
					YoungestChildren = rsMenu("YoungestChildren")
					CategoryID = rsMenu("CategoryID")
					CategoryName = rsMenu("CategoryName")                                                
					Call subMenuMoblie(YoungestChildren,CategoryID,varLink,CategoryName,active)     

				rsMenu.MoveNext()
				loop
%>
                    <li class="nav__item">
                        <form action="/timkiem.html" method="post" name="fTimKiemNhanhMobile" id="fTimKiemNhanhMobile">
                            <div class="search-box-mobile">
                                <input type="text" name="keywordMobile" placeholder="Type to search" class="search-txt-mobile" id="keywordMobile" name="keywordMobile">
                                <button  class="search-btn-mobile">
                                <i class="fas fa-search"></i>
                                </button>
                            </div> 
                        </form>
                    </li>
                    </ul>
                </div>
            </nav>
      <%
    end sub
%>
<% sub subMenu(YoungestChildren,CategoryID,varLink,CategoryName,active)
    CID = trim(Replace(Request.QueryString("cid")," ","+"))
    IDC = Trim(Replace(CID,".html",""))
    IF IDC <> "" And IsNumeric(Trim(Replace(CID,".html",""))) THEN
            if trim(CategoryID) = IDC then 
                active = "active-link" 
            elseif Cint(getColVal("newscategory","ParentCategoryID","categoryid = '"&IDC&"'")) =CInt(CategoryID) then 
                active = "active-link" 
            else 
                active = ""
            end if
   END IF
    if YoungestChildren >0 and CategoryID <> 0 then
%>

<%
            sqlSM= "SELECT  * " &_
                   "FROM      NewsCategory " &_
                   "WHERE     (ParentCategoryID = '"&CategoryID&"') and (CategoryStatus = 2 or CategoryStatus = 4 ) And  CategoryStatus <> 1  ORDER BY CategoryOrder"
          
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1
%><li class="tab-link <%=active %>">
   
    <%
            IF  not rsSM.eof THEN
    %>
	<a href="<%=varLink %>" style="cursor:pointer;"><%=CategoryName %></a>
    <ul class="sub-menu" role="menu">
        <%
            do while not rsSM.eof

   
            call leftSubMenu(rsSM("YoungestChildren"),rsSM("CategoryID"),getLink(rsSM("CategoryID"),"",rsSM("CategoryName")),rsSM("CategoryName"),"")
         
            rsSM.MoveNext()
            loop
        %>
    </ul>
    <% else%>
	 <a href="<%=varLink %>"><%=CategoryName %></a>
     <%END IF %>
</li>
<%
    else
%>
<li class="tab-link <%=active %>"><a href="<%=varLink %>" ><%=CategoryName %></a></li>
<%
    end if
%>
<% end sub %>
<!-- sub-menu ------100% left -->
<% Sub leftSubMenu(YoungestChildren,CategoryID,varLink,CategoryName,active)
    CID = trim(Replace(Request.QueryString("cid")," ","+"))
    IDC = Trim(Replace(CID,".html",""))
    IF IDC <> "" And IsNumeric(Trim(Replace(CID,".html",""))) THEN
            if trim(CategoryID) = IDC then 
                active = "active-link" 
            elseif Cint(getColVal("newscategory","ParentCategoryID","categoryid = '"&IDC&"'")) =CInt(CategoryID) then 
                active = "active-link" 
            else 
                active = ""
            end if
   END IF
    if YoungestChildren >0 and CategoryID <> 0 then
%>
<%
            sqlSM= "SELECT  * " &_
                   "FROM      NewsCategory " &_
                   "WHERE     (ParentCategoryID = '"&CategoryID&"') and (CategoryStatus = 2 or CategoryStatus = 4 ) And  CategoryStatus <> 1  ORDER BY CategoryOrder"
          
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1
%>
<li class="tab-link <%=active %>">
    <%
            IF  not rsSM.eof THEN
    %>
	<a href="#" style="cursor:pointer;"><%=CategoryName %></a>
    <ul class="sub-menu2" role="menu">
        <%
            do while not rsSM.eof
            call leftSubMenu(rsSM("YoungestChildren"),rsSM("CategoryID"),getLink(rsSM("CategoryID"),"",rsSM("CategoryName")),rsSM("CategoryName"),"")
            rsSM.MoveNext()
            loop
        %>
    </ul>
    <% else%>
	 <a href="<%=varLink %>"><%=CategoryName %></a>
     <%END IF %>
</li>
<%
    else
%>
<li class="tab-link <%=active %>"><a href="<%=varLink %>" ><%=CategoryName %></a></li>
<%
    end if
%>

<% end Sub%>

      <!-- sub-menu-moblie bars -->
      <% sub subMenuMoblie(YoungestChildren,CategoryID,varLink,CategoryName,active)



    CID = trim(Replace(Request.QueryString("cid")," ","+"))

    IDC = Trim(Replace(CID,".html",""))
    IF IDC <> "" And IsNumeric(Trim(Replace(CID,".html",""))) THEN
            if trim(CategoryID) = IDC then 
                active = "active-link" 
            elseif Cint(getColVal("newscategory","ParentCategoryID","categoryid = '"&IDC&"'")) =CInt(CategoryID) then 
                active = "active-link" 
            else 
                active = ""
            end if
   END IF
    if YoungestChildren >0 and CategoryID <> 0 then
%>

<%
            sqlSM= "SELECT  * " &_
                   "FROM      NewsCategory " &_
                   "WHERE     (ParentCategoryID = '"&CategoryID&"') and (CategoryStatus = 2 or CategoryStatus = 4 ) And  CategoryStatus <> 1 AND NOT CategoryLoai='9'  ORDER BY CategoryOrder"
          
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1
%>
<li class="nav__item">
   
    <%
            IF  not rsSM.eof THEN
    %>
	<a href="javascript:void(0)" style="cursor:pointer;" class ="nav__link dropdown__link <%=active %>"><%=CategoryName %> <i class="fas fa-chevron-down dropdown__icon"></i></a>
    <ul class="dropdown__menu height" role="menu">
        <%
            do while not rsSM.eof

   
            call subMenuMoblie(rsSM("YoungestChildren"),rsSM("CategoryID"),getLink(rsSM("CategoryID"),"",rsSM("CategoryName")),rsSM("CategoryName"),"")
         
            rsSM.MoveNext()
            loop
        %>
    </ul>
    <% else%>
	 <a href="<%=varLink %>"><%=CategoryName %></a>
     <%END IF %>
</li>
<%
    else
%>
<li class="nav__item <%=active %>"><a href="<%=varLink %>" ><%=CategoryName %></a></li>
<%
    end if
    
%>
<% end sub %>
<%Sub ShowPicture(PictureId,PictureAlign,PictureDirection,NewsId,CatId,title)
  'Nếu NewsId<>0 and CatId<>0: Hiển thị ảnh nhỏ, bấm vào ảnh nhỏ ra chi tiết tin
  'Còn lại: Hiển thị ảnh nhỏ, bấm vào ảnh nhỏ ra ảnh lớn
  
  'PictureDirection: Chỉ có tác dụng khi hiển thị ảnh cùng với Description
  '         =0: width=66
  '         =1: width=100
  '         =2: width=240 (Ảnh nổi bật)
  '         =3: width=46 (Ảnh hiển thị kèm theo tiêu đề tin nổi bật ở bên phải)
  '         =#:Hiển thị đúng kích thước của ảnh nhỏ
  '         =4: width=100 and high= 75
  '         =5:  width=75 and high=100
  Dim rsPic
  set rsPic=server.CreateObject("ADODB.Recordset")
  sql="SELECT * from Picture where PictureId=" & PictureId
  rsPic.Open sql,con,1
  if rsPic.eof then
    rsPic.close
    set rsPic=nothing
    exit Sub
  end if
  LargePictureFileName=rsPic("LargePictureFileName")
  if (LargePictureFileName<>"") and (PictureDirection = 6 or PictureDirection = 7) then
    strNamePic  = rsPic("LargePictureFileName")
  else
    strNamePic  = rsPic("SmallPictureFileName")
  end if
  PictureAuthor = Trim(rsPic("PictureAuthor"))
  PictureCaption  = Trim(rsPic("PictureCaption"))
  Select case PictureDirection
    case 0
      strWidth  =   " width='66' "
    case 1
      strWidth  =   " width='100'"
    case 2
      strWidth  =   " width='240'"
    case 3
      strWidth  =   " width='50'"
    case 4
      strWidth  =   " width='100' height='75' "
    case 5
      strWidth  = " width='75' height='100' "
    case 6
      strWidth  = " width='1000' height='300' "
    case 7
      strWidth  = " width='1000' height='300' "
    case 8
      strWidth  = " width='220' height='125' "
    case 9
      strWidth  = " width='120' "  
    case 10
      strWidth = " width = '100' height = '70'"   
    case else
  End Select  
  ImagePath = NewsImagePath & strNamePic
  if Clng(NewsId)<>0 and Clng(CatId)<>0 then
%>
<!--  <table border="0" cellspacing="0" cellpadding="0"  align="<%=PictureAlign%>" >
        <tr>
           <td > -->
      <a href="/<%=title%>/<%=CatId%>/<%=NewsId%>.html">

      <img  alt="Image" class="w3-left w3-margin-right   " src="<%=ImagePath%>"  <%=strWidth%> >

      </a>
  <!--  </td>
    </tr>
  </table>
 --><%    
  else 'Hiển thị cùng chi tiết tin
'   Response.Write("Giá trị NewsId"&NewsId&"CatId"&CatId)
'   ImagePath=  NewsImagePath&rsPic("SmallPictureFileName")
    iWidth = 0
    iHeight = 0
    if rsPic("SmallPictureFileName")<> "" then
      PathIm  = Path&"\"&rsPic("SmallPictureFileName")
      ImagePath=  NewsImagePath&rsPic("SmallPictureFileName")
'     Response.Write(ImagePath)
      On Error Resume Next
      set myImg = LoadPicture(PathIm)
      On Error Resume Next
      iWidth =  round(myImg.width / 26.4583) 
      On Error Resume Next 
      iHeight = round(myImg.height / 26.4583)-5
'     Response.Write("iWidth="&iWidth&";iHeight="&iHeight)
      if iWidth = 0 or iHeight = 0 then
        iWidth  = 200
        iHeight = 290     
      end if
    else
      ImagePath = "../images/bar9.gif"
      iWidth  = 200
      iHeight = 290
    end if
    'Response.Write("<br> độ dài của ảnh là:"&iWidth&"x"&iHeight)
    sql = "SELECT top 1 NewsCategory.CategoryLoai FROM News INNER JOIN NewsDistribution ON News.NewsID = NewsDistribution.NewsID INNER JOIN NewsCategory ON NewsDistribution.CategoryID = NewsCategory.CategoryID WHERE (News.NewsID = '"&NewsId&"')"
    set rsPic=server.CreateObject("ADODB.Recordset")
    rsPic.Open sql,con,1
    if not rsPic.eof then
      CategoryLoai  = rsPic("CategoryLoai")
    end if
    set rsPic=nothing
    if CategoryLoai = 3 or CategoryLoai=10 or CategoryLoai=7 then
      txtPic  = " width='200' "
      strImage  = "../images/logos.gif"
    else
      strImage  = "../images/bullet.gif"
      txtPic  = " width='"&iWidth&"' height='"&iHeight&"' "
    end if
    
  %>
  
  <table border="0" cellspacing="0" cellpadding="0"  align="<%=PictureAlign%>" >
                <tr>
                  <td  >
            <%
            if LargePictureFileName<>"" then
          %>
          <a href="javascript: openImage('<%=NewsImagePath%><%=LargePictureFileName%>');">
            <img src="<%=ImagePath%>" border="0" <%=txtPic%> >
          </a>
          <%Else
          %>
          <a href="javascript: openImage('<%=ImagePath%>');">
            <img src="<%=ImagePath%>" border="0" <%=txtPic%> >
          </a>
          <%
          End if
          %>
          
          </td>
                </tr>
        <tr><td align="center"><font face="Arial" size="1"><%=PictureCaption%><%if PictureAuthor<>"" then%>&nbsp;Ảnh: <%=PictureAuthor%><%End if%></font></td></tr>
  </table>
  <%end if
  rsPic.close
  set rsPic=nothing
End sub%>




<%sub write_Ads(CatId_,lang,AdsPosition,AdsCount,AdsDirection)
    'AdsPosition=0="Top Banner (450x230)"
    'AdsPosition=1="Banner chuyên mục trong (Width ~ 740)"
    'AdsPosition=2="Banner giữa trang chủ (560x60)"
    'AdsPosition=3="Icon bên trái (width~180)"
    'AdsPosition=4="Icon bên phải (Width ~ 180)"
    'AdsPosition=5="Icon bên phải (Width ~ 580) 
    'AdsDirection=0=Chiều dọc
    '            =1=Chiều ngang
    Varlinks=""

    IF CatId_ <> "" and IsNumeric(CatId_)  THEN 
       
    ELSE
        CatId_ ="1" 
    END  IF

    set rs=Server.CreateObject("ADODB.Recordset")
    if AdsCount>0 then
        sqlAds="SELECT  top " & AdsCount
    else
        sqlAds="SELECT"
    end if
    sqlAds=sqlAds & " Ads_id, Ads_Title, Ads_Link, Ads_ImagesPath, Ads_Type, Ads_width, Ads_height, Ads_Target,Ads_Note,idcolor_tex1,idcolor_tex2, Ads_url FROM V_Ads WHERE (Ads_Position = " & AdsPosition & ") "
    'sqlAds=sqlAds & " Ads_id, Ads_Title, Ads_Link, Ads_ImagesPath, Ads_Type, Ads_width, Ads_height, Ads_Target,Ads_Note,idcolor_tex1,idcolor_tex2, Ads_url FROM    V_Ads " &_
    '       " WHERE (Ads_Position = " & AdsPosition & ") and ((CategoryId=0) or (CategoryId=" & CatId_ & ")"
    strCat=Trim(GetListParentCat(CatId))
    if strCat<>"" then
        ArrCat=Split(" " & strCat & " ")
        for i=1 to UBound(ArrCat)
            if IsNumeric(ArrCat(i)) then
                sqlAds=sqlAds & " or (CategoryId=" & ArrCat(i) & " and Ads_OnlineChildren=1)"
            end if
        next
    end if
    'sqlAds=sqlAds & ") "
    rs.open sqlAds,con,3

%>

<%
    'Ads_Type:  0: GIF,JPG
    '           1: swf
    '           2: txt
    SELECT CASE  0 '--> Clng(rsAds("Ads_Type"))
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------  
        CASE 0

        IF Not rs.EOF THEN    
        itm = 0
        ito = 0

%>
<div class="container-fluid sli-wr">
    <div class="sli-item">
        <div id="carousel-example-generic" class="carousel slide" data-ride="carousel">            
            <!-- Wrapper for slides -->
            <div class="carousel-inner" role="listbox">
            <%
            Do while not rs.eof
            IF itm  = 0 THEN 
               ac = " active "
            ELSE
               ac = " "
            END IF            
            ads_title = Trim(rs("Ads_Title"))
            ads_Note = Trim(rs("Ads_Note"))
            Ads_Link = Trim(rs("Ads_Link"))
            idcolor_tex1 = Trim(rs("idcolor_tex1"))
            idcolor_tex2 = Trim(rs("idcolor_tex2"))
            url_img = "style= 'background:url(/images_upload/"&Trim(rs("Ads_ImagesPath"))&")"&";'" 
        
            idcolor_tex1 = "style= 'color:"&Trim(rs("idcolor_tex1"))&";'" 
            idcolor_tex2 = "style= 'color:"&Trim(rs("idcolor_tex2"))&";'" 
            %>
                <div class="item   <%=ac %>  " <%'=url_img %>>
                    <img src="/images_upload/<%=Trim(rs("Ads_ImagesPath"))%>" alt="<%=ads_title%>">
                    <div style="display:none; overflow: hidden;" class="sli-cap">
                        <div style="display: table-cell; vertical-align: middle;" class="text-center">
                            <h3 style="display:none;" class="cap-title" <%=idcolor_tex1%>><%=ads_title%></h3>
                            <p style="display:none;" class="cap-desc" <%=idcolor_tex2%>><%=ads_Note %></p>
                            <br />
                            <a style="display:none;" href="<%=Ads_Link %>" class="btn-line">Chi tiết</a>
                        </div>
                    </div>
                </div>
            <%  
            itm = itm +1  
            rs.movenext
            Loop
            %>
            </div>
            <!-- Controls -->
            <a class="left carousel-control" href="#carousel-example-generic" role="button" data-slide="prev">
                <span class="fa fa-angle-left glyphicon-chevron-left" aria-hidden="true"></span>
                <span class="sr-only">Previous</span>
            </a>
            <a class="right carousel-control" href="#carousel-example-generic" role="button" data-slide="next">
                <span class="fa fa-angle-right glyphicon-chevron-right" aria-hidden="true"></span>
                <span class="sr-only">Next</span>
            </a>
        </div>
    </div>
</div>
<%    
  END  IF       
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------      
        CASE 1
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------  
            
        CASE 2
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------      
        CASE 9
        ads_title = Trim(rs("Ads_Title"))
        url_img = Trim(rs("Ads_ImagesPath"))
        ads_Note = Trim(rs("Ads_Note"))
        Ads_Link = Trim(rs("Ads_Link"))
%>
<div class="col-md-12" style="clear: both;">
    <a href="<%=Ads_Link %>">
        <img src="<%=url_img %>" class="img-responsive" /></a>
</div>
<%
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------              
        END SELECT
    rs.close
    set rs=nothing
End Sub
%>

<%sub write_Ads2(CatId_,lang,AdsPosition,AdsCount,AdsDirection)
    Varlinks=""

    IF CatId_ <> "" and IsNumeric(CatId_)  THEN 
       
    ELSE
        CatId_ ="1" 
    END  IF

    set rs=Server.CreateObject("ADODB.Recordset")
    if AdsCount>0 then
        sqlAds="SELECT  top " & AdsCount
    else
        sqlAds="SELECT"
    end if
    sqlAds=sqlAds & " Ads_id, Ads_Title, Ads_Link, Ads_ImagesPath, Ads_Type, Ads_width, Ads_height, Ads_Target,Ads_Note,idcolor_tex1,idcolor_tex2, Ads_url FROM V_Ads WHERE (Ads_Position = " & AdsPosition & ") "
    'sqlAds=sqlAds & " Ads_id, Ads_Title, Ads_Link, Ads_ImagesPath, Ads_Type, Ads_width, Ads_height, Ads_Target,Ads_Note,idcolor_tex1,idcolor_tex2, Ads_url FROM    V_Ads " &_
    '       " WHERE (Ads_Position = " & AdsPosition & ") and ((CategoryId=0) or (CategoryId=" & CatId_ & ")"
    strCat=Trim(GetListParentCat(CatId))
    if strCat<>"" then
        ArrCat=Split(" " & strCat & " ")
        for i=1 to UBound(ArrCat)
            if IsNumeric(ArrCat(i)) then
                sqlAds=sqlAds & " or (CategoryId=" & ArrCat(i) & " and Ads_OnlineChildren=1)"
            end if
        next
    end if
    'sqlAds=sqlAds & ") "
	'response.write(sqlAds)
    rs.open sqlAds,con,3

%>

<%
    'Ads_Type:  0: GIF,JPG
    '           1: swf
    '           2: txt
    SELECT CASE  0 '--> Clng(rsAds("Ads_Type"))
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------  
        CASE 0

        IF Not rs.EOF THEN    
        itm = 0
        ito = 0 
%>
<div class="tab-pane  active" id="home">
            <%
            'Dùng vòng lặp check kiểu Ads_Type thiết lập là loại nào
            i=1
            Do while not rs.eof
            IF itm  = 0 THEN 
               ac = " active "
            ELSE
               ac = " "
            END IF            
            ads_title = Trim(rs("Ads_Title"))
            ads_Note = Trim(rs("Ads_Note"))
            Ads_Link = Trim(rs("Ads_Link"))

            Ads_width       =   rs("Ads_width")
            Ads_height      =   rs("Ads_height")
            Ads_ImagesPath  =   rs("Ads_ImagesPath")
            Ads_ImagesPath  =   "/images_upload/"&Ads_ImagesPath

            if i=1 then
                Response.write("<div class='owl-carousel owl-theme  z-index' id='carousel-1'>")
            end if
			response.write ("<a href='" &link& "' "&itarget&">")
		            response.write "<img src='" & Ads_ImagesPath & "' alt='" & rs("Ads_Title") & "' title='" & rs("Ads_Title") & "'>"
			response.write ("</a>") 
            if i=rs.recordcount  then
                Response.write "</div>"
             
         
            end if 

            i=i+1
            itm = itm +1  
            rs.movenext
            Loop
            %>
    <script type="text/javascript" lang="javasript">
        var owl = $("#home .owl-carousel");
        owl.owlCarousel({
            items: 1,
            loop: true,
            nav: false,
            margin: 10,
            autoplay: true,
            autoplayTimeout: 2500,
            autoplayHoverPause: true,
            lazyLoad: true,
            dots: true,
        });
        window.addEventListener("load", () => {
            const owlCarousel = document.getElementById("home");

            owlCarousel.scrollIntoView({ behavior: "smooth", block: "end", inline: "center" });


        });
    </script>
</div>
   

     
<%    
  END  IF       
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------      
        CASE 1
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------  
            
        CASE 2
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------      
        CASE 9
        ads_title = Trim(rs("Ads_Title"))
        url_img = Trim(rs("Ads_ImagesPath"))
        ads_Note = Trim(rs("Ads_Note"))
        Ads_Link = Trim(rs("Ads_Link"))
%>
<div class="col-md-12" style="clear: both;">
    <a href="<%=Ads_Link %>">
        <img src="<%=url_img %>" class="img-responsive" /></a>
</div>
<%
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------              
        END SELECT
    rs.close
    set rs=nothing
End Sub
%>


<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- %>

<%
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
Sub  Ineter_F(cid_,idindex) %> <!--=(CategoryID,1)-->     

<%IF idindex <> "" And idindex = 4 THEN  ' Dịch vụ 
    
    sql = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')   And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'CateHome = -1 :đặc biệt
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1
    
    Response.Write "<br />"   
    if not rsi.eof then   
    CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
        
%>
<div class="container">
    <hr class="Hr-Title" />
    <div class="form-group" style="position: relative; height: 36px;">
        <div class="in-title col-lg-3 col-md-3 col-sm-6 col-xs-6">
            <h3 class="H-Title"><%=CName %></h3>
        </div>
    </div>

    <div class="row sv-line">
        <%
            do while Not rsi.EOF
                tem_ = 0
                Title   = Trim(rsi("Title"))
                desc    = Trim(rsi("Author"))

                if  Len(desc)  < 95  then 
                    desc_ = desc
                else
                    desc_ = Left(desc,95)&"..."
                end if

                desc_short    = Trim(rsi("DecsBannerImage"))        
                cateID  = Trim(rsi("CategoryID"))
                NewsID  = Trim(rsi("NewsID"))
                nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
                linkuri = func_uri(cateID,NewsID,Title)                         
        %>

        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" alt=<%=Title%> />
                    <p class="sv-title"><%=Title %></p>
                </a>
            </div>
            <p class="sv-desc"><%=desc_ %></p>
        </div>
        <%
            rsi.MoveNext
            Loop
        %>
    </div>

</div>
<%
    set rsi = nothing
    end if 'end if dich vu
%>

<%ELSEIF idindex <> "" And idindex = 3 THEN  ' Dịch vụ  2   
    sql = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')   And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'CateHome = -1 :đặc biệt
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1 
    Response.Write "<br />"   
    if not rsi.eof then   
     CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
        
%>
<div class="container">
    <hr class="Hr-Title" />
    <div class="form-group" style="position: relative; height: 36px;">
        <div class="in-title col-lg-3 col-md-3 col-sm-6 col-xs-6">
            <h3 class="H-Title"><%=CName %></h3>
        </div>
    </div>

    <div class="row sv-line">
        <%
            do while Not rsi.EOF
            tem_ = 0
            Title   = Trim(rsi("Title"))
            desc    = Trim(rsi("Author"))

            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if

            desc_short    = Trim(rsi("DecsBannerImage"))                
            cateID  = Trim(rsi("CategoryID"))
            NewsID  = Trim(rsi("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
            url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)                   
        %>
        <div class="col-md-6 sv-item">
            <div class="sv-items" <%=url_img %>>
                <div style="display: table; overflow: hidden;" class="sv2-w">

                    <div style="display: table-cell; vertical-align: middle;" class="text-center sv2-title">
                        <a href="<%=linkuri %>"><%=Title %></a>
                    </div>
                </div>
            </div>
        </div>
        <%
            rsi.MoveNext
            Loop
        %>
    </div>
</div>
<%
    set rsi = nothing
    end if 'end if dich vu 
%>

<%ELSEIF idindex <> "" And idindex = 2 THEN  ' Ineter_F(cid_,idindex) ' =(CategoryID,4) ở đây gồm các tìn từ thanh bars ||| đây là dịch vụ
    sql = "SELECT  top 16 * FROM V_News WHERE CategoryID = '"&cid_&"' or ParentCategoryID='"& cid_ &"' and  status = '4'  And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'cid id ở đây là catergoryID
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1 
    Response.Write "<br />"   
    if not rsi.eof then   
    CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
    link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html"        
%>
<div class="container" style="margin-bottom:20px;"> 
    <% if trim(CName) = "Góc chia sẻ" then%>
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title col-lg-12 col-md-12 col-sm-6 col-xs-12">
            <a href="<%=link_cate %>" class="d-flex"><h3 class="H-Title" style="margin:auto"><%=CName %></h3></a>        
        </div>       
    </div>
    <div class="text-center px-3 pb-2"><i>Phản hồi của những khách hàng đã và đang sử dụng sản phẩm trong suốt những năm qua</i></div>
    <% else %>
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title col-12 col-md-6">
            <a href="<%=link_cate %>"><h3 class="H-Title"><%=CName %></h3></a>        
        </div>       
    </div>
    <% End if %>
        <%
            item=1
            Do while Not rsi.EOF
            tem_ = 0
            Title   = Trim(rsi("Title"))
            desc    = Trim(rsi("Author"))
            desc_short    = Trim(rsi("DecsBannerImage"))
            if rsi("PriceNet") <> ""  then
                originPrice = Clng(Trim(rsi("Price")))
                discountpercent = Clng(trim(rsi("PriceNet"))) 
                Price = Clng(rsi("Price")) - ((discountpercent * Clng(rsi("Price"))) / 100 )
                Price = Dis_str_money(Price)
            else
                Price = Dis_str_money(Trim(rsi("Price")))
            End IF 
            Unit = Trim(rsi("Unit"))  
            if  Len(desc_short)  > 100  then 
                desc_short = Left(desc_short,100)&"..."
            end if

            cateID  = Trim(rsi("CategoryID"))
            NewsID  = Trim(rsi("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
            url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)
            
            if item=1 then
                Response.write "<div class='p-5 rounded' style='background-color: white;'><div class='index-group owl-carousel owl-theme '>"
            end if                    
       %>
            <div class="item product-under w3-border w3-round-xxlarge text-center"> 
                 <figure class="product-image"> 
                <a href="<%=linkuri %>"><img src="/images_upload/<%=nimg %>" class="img-responsive mx-auto d-block"></a>
                    <div class="product-over">
                                        <button
                                            class="btn btn-small addToCart"
                                            data-product-id="<%=NewsID%>"
                                        >
                                            <i class="fas fa-cart-plus"></i>Add
                                            to cart
                                        </button>
                                        <a
                                            href="<%=linkuri %>"
                                            class="btn btn-small"
                                            >Đọc thêm</a
                                        >
                    </div>
                </figure>
                <div class="news-content w3-padding">
                    <p class="news-title">
                    <a href="<%=linkuri %>" class="w3-text-black productName"><b><%=Title %> </b></a>
                </p>
                    <span class="stars"></span>
                    <div class="news-desc "><%=desc_short %></div>
                    <% if Price <> "" and Price <> 0 then response.Write("<span class='news-desc w3-text-green w3-large priceValue'>"&Price&"</span>"&"<span class='News-unit w3-text-green'>"&Unit&"</span>")%>
                <!--<div class="news_detail"><a href="<%=linkuri %>">Xem thêm</a></div>-->
                 </div>
            </div>
        <%
            if item = rsi.recordcount then 
                response.Write("</div></div>") '<!---/.index-group--->
            end if


            item=item+1
            rsi.MoveNext
            Loop
        %>
        <script>
            var owl = $(".index-group");
            //console.log(owl);
            owl.owlCarousel({
                items: 4,
                loop: true,
                margin: 10,
                autoplay: true,
                autoplayTimeout: 2000,
                autoplayHoverPause: true,
                nav: true,
                responsive: {
                    0: {
                        items: 1
                    },
                    600: {
                        items: 2
                    },
                    1000: {
                        items: 4
                    }
                }
            });
        </script>
</div>

<%
    set rsi = nothing
    end if 'end if dich vu %>

<%ELSEIF idindex <> "" And idindex = 8 THEN  ' Đối tác
    
    sql = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')   And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'CateHome = -1 :đặc biệt
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1 
    Response.Write "<br />"   
    if not rsi.eof then   
        tem_ = 0
        Title   = Trim(rsi("Title"))
        desc    = Trim(rsi("Author"))

        if  Len(desc)  < 95  then 
            desc_ = desc
        else
            desc_ = Left(desc,95)&"..."
        end if


        desc_short    = Trim(rsi("DecsBannerImage"))
            
        cateID  = Trim(rsi("CategoryID"))
        NewsID  = Trim(rsi("NewsID"))
        nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
        url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
        linkuri = func_uri(cateID,NewsID,Title)         
        CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
%>
<div class="container">
    <hr class="Hr-Title" />
    <div class="form-group" style="position: relative; height: 36px;">
        <div class="in-title col-lg-3 col-md-3 col-sm-6 col-xs-6">
            <h3 class="H-Title"><%=CName %></h3>
        </div>
    </div>

    <div class="index-news">
        <%
            do while Not rsi.EOF
        %>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <div class="col-md-3 news-item">
            <div class="news-items">
                <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                <p class="news-title"><b><%=Title %> </b></p>
                <p class="news-desc"><%=desc %></p>
            </div>
        </div>
        <%
            rsi.MoveNext
            Loop
        %>
    </div>
</div>
<%
    set rsi = nothing
    end if 'end if dich vu 
%>
<%       
    END IF      
    end sub
%>


<%'------------------------------------------------------------------------------------------------------------------------------------------------------- %>
<%sub Fs_Paner(cid_,LgID) 
    sqlPaner = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')  Order By   LastEditedDate DESC "  
    set RsPaner = server.CreateObject("ADODB.RECORDSET")
    RsPaner.open sqlPaner,con,1   
    if not RsPaner.eof then   
       CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
       CNote =  getColVal("newscategory","categoryNote","categoryid = '"&cid_&"'")
        
       
%>
<div class="container pdb-20">
    <hr class="Hr-Title">
    <div style="float: left; background: #75bb1f;" class="in-title">
        <h3 class="H-Title"><%=CName %></h3>
    </div>
    &nbsp;
    <br />

    <p class="sv-cnote clear"><%=CNote %></p>
    <p class="intro-title  clear text-center">
        CÁC CÔNG TRÌNH
        <br />
        ĐANG CUNG CẤP DỊCH VỤ VỆ SINH THƯỜNG XUYÊN
    </p>

</div>
<div class="container pdb-80">
    <div class="row sv-line">
        <%
            do while Not RsPaner.EOF
            Title   = Trim(RsPaner("Title"))
            desc    = Trim(RsPaner("Author"))

            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if
            desc_short    = Trim(RsPaner("DecsBannerImage"))              
            cateID  = Trim(RsPaner("CategoryID"))
            NewsID  = Trim(RsPaner("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&RsPaner("PictureId")&"'")   
            url_img = "style= 'background:url(/images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)         

        %>
        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                    <p class="Pal-Title"><%=Title %></p>
                </a>
            </div>
        </div>
        <%
            RsPaner.MoveNext
            Loop
        %>
    </div>
</div>
<%
    set RsPaner = nothing
    end if 'end if dich vu 
    End sub
%>
<%'------------------------------------------------------------------------------------------------------------------------------------------------------- %>

<%sub Tindacbiet() %>

<div class="container-fluid">

    <div class="container">
        <div class="news-wbd">


            <h2 class="News-Title2">TẠI SAO CHỌN GICO LÀM NHÀ CUNG CẤP DỊCH VỤ?</h2>
            <p class="news-desc" style="text-align: center;"><b>Lựa chọn sử dụng các dịch vụ của GICO, Quý khách hàng chắc chắn sẽ được hưởng các lợi ích sau:</b></p>
            <p class="news-desc">
                <img src="/images/icon/tick-1.png" />
                Chất lượng cao, chi phí hợp lý.<br />
                <img src="/images/icon/tick-1.png" />
                Được tư vấn và cung cấp giải pháp làm sạch, bảo trì dựa trên nhu cầu của khách hàng và hiện trạng của tài sản.<br />
                <img src="/images/icon/tick-1.png" />
                Được sử dụng đội ngũ nhân viên chuyên nghiệp, lành nghề của công ty hàng đầu ở Việt Nam trong lĩnh vực vệ sinh công nghiệp và bảo trì bất động sản..<br />
                <img src="/images/icon/tick-1.png" />
                Không phải đầu tư máy móc, thiết bị làm sạch chuyên dùng vì vậy tiết kiệm được chi phí đầu tư<br />
                <img src="/images/icon/tick-1.png" />
                Không những duy trì được giá trị tài sản, dịch vụ vệ sinh chuyên nghiệp còn mang lại những giá trị tăng như: môi trường làm việc sạch đẹp cho cán bộ công nhân viên, khách tham quan và mang lại ấn tượng tốt đối với khách hàng<br />
                <img src="/images/icon/tick-1.png" />
                Không phải trực tiếp quản lý, điều hành công nhân vệ sinh. Mọi vấn đề phát sinh đều đã có GICO đứng ra chịu trách nhiệm giải quyết, nhờ vậy tránh được các rủi ro (nếu có)<br />
                <img src="/images/icon/tick-1.png" />
                Chia sẻ trách nhiệm trong công tác quản lý, khách hàng có thể toàn tâm toàn ý tập trung vào công việc kinh doanh của mình<br />
            </p>
        </div>
    </div>

</div>
<%end sub %>
<%
   '09-09-2015 TROJAN  09-09-2015 TROJAN  09-09-2015  TROJAN 09-09-2015 TROJAN  
Sub Fs_Faq()
%>
<div class="container pb-40">
    <hr class="Hr-Title">
    <h3 class="H-Title">Các câu hỏi thường gặp</h3>
    <p class="br-line"></p>
    <br>
    &nbsp;
    <br>
</div>





<div class="container" id="Faq_">
    <div class="col-lg-12">

        <%  
    sqlFaq="SELECT * from Y_KIEN where show <> '0'"
    Set rsF=Server.CreateObject("ADODB.Recordset")
    rsF.open sqlFaq,con,3   

    do while not rsF.eof
        id=Trim(rsF("id"))
        CMND=Trim(rsF("CMND"))
        tel=rsF("tel")
        hovaten=Trim(rsF("hovaten"))
        tieude= Trim(rsF("tieude"))
        noidung=Trim(rsF("noidung"))
        faq=rsF("faq")
        show=rsF("show")
        ngaytao=rsF("ngaytao")
        Traloi  =rsF("Traloi")


            
        %>

        <div class="col-md-6">
            <div class="panel-group" id="accordion<%=stt %>" role="tablist" aria-multiselectable="true" style="width: 90%; margin: auto auto">
                <div class="panel panel-default Faq-tab">
                    <div class="panel-heading" role="tab" id="headingOne<%=stt %>">
                        
                        <a class="faq-collapse collapsed" role="button" data-toggle="collapse" data-parent=".accordion<%=stt %>" id="key_<%=stt %>" href="#colla<%=stt %>" aria-expanded="true" aria-controls="collapseOne">
                            <span class=" w3-text-amber w3-large"> 
                                <%=hovaten %>  :</span> 
                            <span class="panel-title">
                                <%=tieude %>  <i class="fa fa-minus faq-k collapse in" aria-hidden="true"></i></span>
                        </a>

                    </div>


                    <!--<div class="panel-body Faq-text Faq_Acdesc collapse in key_<%=stt %>">
                        <%=Left(Traloi,70) %>...
                    </div>-->
                    <div id="colla<%=stt %>" class="panel-collapse  collapse  " role="tabpanel" aria-labelledby="headingOne<%=stt %>">
                        <div class="panel-body Faq-text collapse in ">

                            <b>Câu hỏi: </b><%=noidung %>
                        </div>
                        <div class="panel-body Faq-text collapse in ">

                            <b>Admin trả lời : </b><%=Traloi %>
                        </div>


                    </div>
                </div>
            </div>


        </div>
        
        <%

            stt = stt+1
    rsF.movenext
    Loop
    rsF.close
    set rsF=nothing
        %>
    </div>
</div>
      <br />

<script type="text/javascript">
    $(document).ready(function () {
        $(".faq-collapse").click(function () {
            if ($(this).hasClass("collapsed")) {
                $("div").find("." + $(this).attr('id') + "").removeClass("in");
                $(this).find("i").removeClass("fa fa-minus");
                $(this).find("i").addClass("fa fa-plus");
            }
            else {
                $("div").find("." + $(this).attr('id') + "").addClass("in");
                $(this).find("i").removeClass("fa fa-plus");
                $(this).find("i").addClass("fa fa-minus");

            }

        });
    });
</script>
<%  
end sub
%>


<% 
Sub Fs_Contact(Cate_,LgID)

    CName       =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")

%>

<div class=" container-fluid">
    <div class="container">
        <hr class="Hr-Title">
        <div class="form-group" style="position: relative; height: 36px;">
            <div style="float: left; background: #75bb1f;" class="in-title">
                <h3 class="H-Title"><%=CName %></h3>
            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i><%=CName %> </i></div>
        </div>

        <h2 class="intro-title text-center" style="display:none;">Bản đồ
            <br />
            <label class="lh-slg"><%=company %></label>
        </h2>
        <div class="maps-gool pdb-40" style="min-height: 250px;display:none;">
            <iframe src="https://www.google.com/maps/d/embed?mid=1FmOeFQS7LFzhV1Xt0KQ9rNnbUvo" width="100%" height="250" frameborder="0" style="border: 0" allowfullscreen></iframe>
        </div>
        <p class="txt-nomal text-center">
            xin chân thành cảm ơn Quý khách đã quan tâm đến dịch vụ của chúng tôi.
            <br />
            Mọi thông tin vui lòng liên hệ trực tiếp với chúng tôi theo thông tin sau:
            <br />
            &nbsp;
            <br />
        </p>

        <div class="form-group">
            <div class="col-md-7">
                <h4 class="lh-slg up-care">
                    <%=company %>
                </h4>

                <label>Địa chỉ: <%=Address %></label><br />
                <label>Hotline:<b style="color: #F00;"><%=Hotline %></b></label><br />
                <label>Fax: <%=Fax %></label><br />
                <label>Email:<%=Email %></label><br />
                <label>Website:<%=Website %></label><br />
            </div>
            <div class="col-md-5">
                <h4 class="lh-slg up-care">Thông Tin Liên Hệ
                </h4>
                <form name="Fremail" id="Fremail" method="post" class="">

                    <div class="form-group">
                        <input type="text" class="form-control" id="F_Name" name="F_Name" placeholder="Họ tên" maxlength="50">
                    </div>
                    <div class="form-group">
                        <input type="email" class="form-control" id="F_Email" name="F_Email" placeholder="Email" maxlength="150">
                    </div>
                    <div class="form-group">
                        <input type="tel" class="form-control" id="F_Title" name="F_Title" placeholder="Tiêu đề" maxlength="10">
                    </div>
                    <div class="form-group">
                        <textarea class=" form-control lh-aria" id="F_Content" name="F_Content" placeholder="Nội dung"></textarea>
                    </div>
                    <div class="form-group text-center">
                        <button id="btnSubmit" type="button" class="btn e-btn">Gửi</button>
                        <button id="" type="reset" class="btn e-btn">Làm mới</button>
                        <br />
                        &nbsp;
                        <br />
                    </div>
                </form>
            </div>
        </div>



    </div>
</div>







<script type="text/javascript">
    $("#btnSubmit").click(function () {

        if ($('#F_Name').val() == '') {
            $('#F_Name').focus();
            swal("BQT", "Xin vui lòng nhập họ tên.");
        }
        else if ($('#F_Email').val() == '') {
            $('#F_Email').focus();
            swal("BQT", "Xin vui lòng nhập email.");
        }
        else if (!isEmail($('#F_Email').val())) {
            $('#F_Email').focus();
            swal("BQT", "Sai định dạng email.vd: abc@gmail.com");
        }
        else if ($('#F_Title').val() == '') {
            $('#F_Title').focus();
            swal("BQT", "Vui lòng nhập tiêu đề.");
        }
        else if ($('#F_Content').val() == '') {
            $('#F_Content').focus();
            swal("BQT", "Vui lòng nhập nội dung.");
        }
        else {
            Func_resEmail('send-contact', '0');
        }

    });

    function isEmail(email) {
        var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
        return regex.test(email);
    }

</script>
<%end sub %>

<%Function GetListParentCatNameOfCatId2(CatId,NewsID)
    ''Get Tree List Name Of CategoryId of Inpute Category.
    'Result is a string of CategoryId's separated by spacebar, include Input Category
    Dim i,strArrValue
    i=0
    str_link=""
    Dim rs1
    set rs1=Server.CreateObject("ADODB.Recordset")

      PCatId=CatId

      str_link="<div class='breadcrumbs-wrap'>"  

      Do while PCatId<>0
        i=i+1
        sql_GetListParentCat="select CategoryId,ParentCategoryId,CategoryName from NewsCategory where CategoryId=" & PCatId
        ' response.write sql_GetListParentCat
        rs1.open sql_GetListParentCat,con,1
            PCatId      =   Cint(rs1("ParentCategoryId"))
            CategoryID  =   Cint(rs1("CategoryId"))
            varLink     =   getLink(rs1("CategoryId"),"",rs1("CategoryName"))
            if i=1 then
                
               if NewsID>0 then
                   'iTitle=GetValueSQL("SELECT Title FROM News WHERE NewsID='"&NewsID&"'","Title")
                   'iTitle=ucase(mid(iTitle,1,1))&lcase(mid(iTitle,2))
                   'link_h1="<li><h1 class='page-title'>"&iTitle&"</h1></li>"
                   'strArrValue="<li><a href='"&varLink&"'>"&rs1("CategoryName")&"</a></li>"                 
                    link_h1="<li><h1 class='page-title'>"&rs1("CategoryName")&"</h1></li>"
                    strArrValue="<li><a href='"&varLink&"'>"&rs1("CategoryName")&"</a></li>"
               else
                    link_h1="<li><h1 class='page-title'>"&rs1("CategoryName")&"</h1></li>"
                strArrValue="<li>"&rs1("CategoryName")&"</li>"
               end if
            else
                if PCatId<> 0 then
                    strArrValue="<li><a href='"&varLink&"'>"&rs1("CategoryName")&"</a></li>"& strArrValue
                else
                    strArrValue="<li><a href='"&varLink&"'>"&rs1("CategoryName")&"</a></li>"& strArrValue  
                end if
            end if
        rs1.close
      Loop  
    str_link=str_link&"<ul class='breadcrumb'>"&link_h1&"<li class='home'><a href='/' title='trở về trang chủ'>Trang chủ</a></li>"&strArrValue+"</ul></div><!---/.breadcrumbs-wrap--->"
    GetListParentCatNameOfCatId2=str_link
End Function%>

<%
Sub Fs_CateAnything(Cate_,Lgid_)
    sqlif = "SELECT * FROM V_News WHERE   status = '4' And  CategoryID = '"&Cate_&"'  And (url_video IS NULL or url_video ='')   And LanguageID='"&Lgid_&"' Order by LastEditedDate DESC "    
    set rsif = server.CreateObject("ADODB.RECORDSET")
    rsif.open sqlif,con,1

    If Not rsif.EOF Then 'IF 1
    if rsif.RecordCount  > 1 then 
     CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
<div class=" container-fluid">
    <div class="container pdb-40">
        <%=GetListParentCatNameOfCatId2(Cate_,0) %>
        <%Call NewsLater() %>        
        <div class="col-md-8">
            <div class="form-group" style="position: relative">
                <p class="H-visible">&nbsp&nbsp&nbsp</p>          
            </div>  
        <div class="news_category">     
        <%  
    Response.write(sqlif)

        item=1
        do while not rsif.eof 
            Title       =   Trim(rsif("Title"))
            Title       =   ucase(mid(Title,1,1))+lcase(mid(Title,2))
            desc_short  =   Trim(rsif("DecsBannerImage"))                           
            if  Len(desc_short)  > 130  then 
                desc_short = Left(desc_short,130)&"..."
            end if
            LastEditedDate  =   rsif("LastEditedDate")
            nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsif("PictureId")&"'")   
            linkuri = func_uri(rsif("CategoryID"),rsif("NewsID"),rsif("Title"))
            if item=1 then
                Response.write "<div class='news_category-item'>"
            end if                     
        %>        
        <div class="news_desc">
            <a href="<%=linkuri %>">
                <img class="img-responsive Av-news3" src="/images_upload/<%=nimg %>" alt=<%=Title%> />
            </a>
            <p class="news3">
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <span class="text-justify add-text"><%=desc_short %></span>
            </p>           
            <hr>            
        </div>             
        <%       
            if item mod 4=0 and item>1 then                      
                Response.write "</div><!---/.news_category-item--->"
                if item < rsif.recordcount then
                    Response.write "<div class='news_category-item'>"
                end if
            end if  
             
            if item = rsif.recordcount then
                Response.write "</div><!---/.news_category-item--->"                
            end if        
            item=item+1
            rsif.MoveNext
        Loop
        %>
        </div><!---/.news_category--->
        </div><!---/.col-md-8--->        
    </div>
</div>
<%
    end if ''IF 2
    End If 
    Call Fs_CateOther()   
End Sub
%>

<%
Sub Fs_CateNews(Cate_,Lgid_)
    sqlif = "SELECT * FROM V_News WHERE   status = '4' And  CategoryID = '"&Cate_&"'  or ParentCategoryID ="&Cate_&"  And LanguageID='"&Lgid_&"' Order by LastEditedDate DESC "    
    set rsif = server.CreateObject("ADODB.RECORDSET")
    rsif.PageSize = 10
    rsif.open sqlif,con,1
    IF Not rsif.EOF   THEN 'IF 1
    IF rsif.RecordCount  > 0 And rsif.RecordCount  < 2 THEN
        Title       = Trim(rsif("Title"))
        desc        = Trim(rsif("DecsBannerImage"))
        FTitle      = Trim(rsif("Description")) 
        Ncontent    = Trim(rsif("Body")) 
        cateID      = Trim(rsif("CategoryID"))
        NewsID      = Trim(rsif("NewsID"))
        nimg        =  getColVal("Picture","SmallPictureFileName","PictureId = '"&rsif("PictureId")&"'")    
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
    <div class="container pdb-40">
        <div class="form-group">
            <div class="in-title css_relative">
                <h3 class="H-Title"><%=CName %></h3>
            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i>&nbsp; <%=CName %> </i></div>
        </div>
        <h1 class="intro-title"><%=Title %></h1>
		<%if FTitle<>"" then%>
        <div class="intro-phu text-center"><%=FTitle %></div>
		<%end if %>
        <div class="intro-body"><%=Ncontent %></div>
    </div>
<%
    ELSEIF rsif.RecordCount  > 1 THEN 
     CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
<div class=" container pdb-40">
        <div class="form-group">
            <div class="in-title css_relative">
                <h3 class="H-Title"><%=CName %></h3>
            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i>&nbsp; <%=CName %> </i></div>
        </div>
        <div class="row justify-content-center">
            <%  
            if request.Querystring("page")<>"" then
			    page=Clng(request.Querystring("page"))
		    else
			    page=1
		    end if

		      rsif.AbsolutePage = CLng(page)
            j=0
            i = 1
    do while not rsif.eof and j<rsif.pagesize 
        Title   = Trim(rsif("Title"))
        desc    = Trim(rsif("DecsBannerImage"))
        FTitle    = Trim(rsif("Description")) 
        Ncontent    = Trim(rsif("Body")) 
        cateID  = Trim(rsif("CategoryID"))
        NewsID  = Trim(rsif("NewsID"))
        nimg    =      getColVal("Picture","SmallPictureFileName","PictureId = '"&rsif("PictureId")&"'")    
        linkuri = func_uri(cateID,NewsID,Title)              
        CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
        %>
        <div class=" col-12 col-md-5 border-bottom m-3"> <!--padding-add-1 -->
            <div class="row ">
                <div class="col-4 p-0">
                    <a href="<%=linkuri %>">
                        <img src="/images_upload/<%=nimg %>" class=" img-responsive news-cate-img" alt="<%=Title%>" />
                    </a>
                </div>
                <div class="col-8">
                    <div class="news-cat-des">
                        <a href="<%=linkuri %>" class="sl-news-Tl"><b><%=Title %></b></a>
                         <%
                            if  Len(desc)  > 200 then 
                                desc = mid(desc_short,1,200)&"..."
                            end if
                         %>
                        <span class="text-justify add-text"><%=desc %></span>
                    </div>
                </div>
            </div>
            <p class="text-right">
                    <a href="<%=linkuri %>" class="cews-cat-dls">Xem tiếp...</a>
            </p>
        </div>
        
        <%     
            i=i+1
            j=j+1
    rsif.MoveNext
    Loop

        pagecount=rsif.pagecount
        pageperbook=10
        %>
    </div>
    <%Call phantrang(page,pagecount,pageperbook) %>
  </div>
<%
    End if
    End if
end sub 
%>

<% Function getUrl(CateID)%> 
  <%
     sql = "SELECT * FROM NewsCategory WHERE  categoryID="&CateID
    'response.write sql
    set rsMenu = Server.CreateObject("ADODB.Recordset")
    rsMenu.open sql,con,1
    If not rsMenu.eof Then 
      url= "/"&Replace(Uni2NONE(rsMenu("CategoryName"))," ","")&"/"&rsMenu("CategoryId")&".html"

    end if  
    set rsMenu=nothing
getUrl= url
        %>
<%End Function %>

<%Sub phantrang(page,pagecount,pageperbook)%>
<%
Response.write "<ul class='pagination'>"
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
    if request.ServerVariables("QUERY_STRING")<>"" then ' trả ra giá trị đằng sau dấu hỏi 
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
    'response.write(request.ServerVariables("SCRIPT_NAME")) ' nó sẽ trả ra 1 virtual path chính là tên file 
    'Bản ghi first và prepage
     if  page-Clng((pageperbook-1)/2) > 1 then
          Response.Write "<li class='page-item'><a class='page-link' href=""" & sURL & """>←</a>"
          Response.Write "<li class='page-item'><a class='page-link' href=""" & sURL & "page=" & page-1 & """>«</a>"
    end if

    for i=minpage to maxpage
		if i<>page then
            if i=1 then
                Response.Write "<li class='page-item'><a class='page-link' href=""" & sURL & """>" & i & "</a></li>"
            else
                Response.Write "<li class='page-item'><a class='page-link' href=""" & sURL & "page=" & i & """>" & i & "</a></li>"
            end if               
		else
			Response.Write "<li class='page-item active'><span class='page-link'>" & i & "</span></li>"
			next_page=i+1
		end if
	Next
            
    'Bản ghi last và nextpage
    if page+Clng((pageperbook-1)/2) < pagecount  then
			Response.Write "<li class='next page-item'><a class='page-link' href=""" & sURL & "page=" & next_page & """>»</a>"
            Response.Write "<li class='last page-item'><a class='page-link' href=""" & sURL & "page=" & pagecount & """>→</a>"
    end if

Response.write "</ul>"
End sub%>

<%'------------------------------------------------------------------------------------------------------------------------------------ %>
<%
Sub Fs_NewsDetail(NewsID) 
    sqlN = "SELECT * FROM V_News WHERE   status = '4' And  NewsID = '"&NewsID&"'  And (url_video IS NULL or url_video ='')"
    set rsN = Server.CreateObject("ADODB.Recordset")
    rsN.open sqlN,con,1
    If not rsN.eof then
        Title       = Trim(rsN("Title"))
        desc        = Trim(rsN("Author"))
        FTitle      = Trim(rsN("Description")) 
        Ncontent    = Trim(rsN("Body")) 
        cateID      = Trim(rsN("CategoryID"))
        NewsID      = Trim(rsN("NewsID"))    
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&cateID&"'")
%>

<div class="container-fluid product_content">
    <div class="container">
        <%=GetListParentCatNameOfCatId2(cateID,NewsID)%>
        <h1 class="intro-title text-center"><%=Title %></h1>
        <div class="intro-body"><%=Ncontent %></div>
    </div><!--/.container--->
</div>
<%
    END IF 'not eof
    rsN.Close
%>
<%End Sub %>

<%'------------------------------------------------------------------------------------------------------------------------------------ %>
<%
Sub Fs_NewsInvolve(CatID,NewsID)
%>
<%
sqlN = "SELECT Top 12 * FROM V_News WHERE status = '4' And CategoryID = '"&CatID&"' And NewsID <> '"&NewsID&"' Order By  LastEditedDate DESC "        
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sqlN,con,1    
    if not rsi.eof then  
%>
<div class=" container-fluid pdb-40">
    <div class="container">
        <div class="col-lg-12 col-md-12">
            <div class="form-group w3-padding-right">
            <div class="in-title">
                <h3 class="H-Title">Tin cùng chuyên mục</h3>            
            </div>       
            </div>
        </div>
        <div class="col-lg-12 col-md-12">     
        <%       
        Response.write "<div class='col-md-12 news_category_other'>" 
        item=1
        ihr=0
        do while Not rsi.EOF
            Title       =   Trim(rsi("Title"))
            LastEditedDate  =   rsi("LastEditedDate")
            nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")             
            linkuri = func_uri(rsi("CategoryID"),rsi("NewsID"),rsi("Title"))
            
            if item=1 then
                Response.write "<div class='news_item '>"
            end if   
        %>     
        <div class="Item-news w3-padding-right col-md-4">
            <a href="<%=linkuri %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
            </p>
            <%if (item=ihr*6+1 or item=ihr*6+2 or item=ihr*6+3) then  %>
            <hr />
            <%end if %>
        </div>
        <%
            if item mod 6=0 and item>1 then    
                ihr=ihr+1                  
                Response.write "</div><!---/.news_item--->"
                if item < rsi.recordcount then
                    Response.write "<div class='news_item col-md-12'>"
                end if
            end if  
             
            if item = rsi.recordcount then
                Response.write "</div><!---/.news-item--->"                
            end if        
            item=item+1
        rsi.MoveNext
        Loop
            Response.write "</div><!---/.col-md-12--->"
        %>
        </div><!--col-md-12-->
    </div>
</div>
    <%end if%>
<%End Sub %>

<%
Sub Fs_CateIntro(Cate_,Lgid_)
    sqlif = "SELECT * FROM V_News WHERE   status = '4' And  CategoryID = '"&Cate_&"'  And (url_video IS NULL or url_video ='')   And LanguageID='"&Lgid_&"' Order by LastEditedDate DESC "    
    set rsif = server.CreateObject("ADODB.RECORDSET")
    rsif.open sqlif,con,1
    IF Not rsif.EOF   THEN 'IF 1
    IF rsif.RecordCount  > 0 And rsif.RecordCount  < 2 THEN
        Title       = Trim(rsif("Title"))
        desc        = Trim(rsif("Author"))
        FTitle      = Trim(rsif("Description")) 
        Ncontent    = Trim(rsif("Body")) 
        cateID      = Trim(rsif("CategoryID"))
        NewsID      = Trim(rsif("NewsID"))
        'nimg        =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsif("PictureId")&"'")    
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
<div class=" container-fluid">
    <div class="container">
        <hr class="Hr-Title">
        <div class="form-group" style="position: relative; height: 36px;">
            <div style="float: left; background: #75bb1f;" class="in-title">
                <h3 class="H-Title"><%=CName %></h3>

            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i><%=CName %> </i></div>
        </div>
        <h1 class="intro-title"><%=Title %></h1>
		<%if FTitle<>"" then%>
        <div class="intro-phu text-center"><%=FTitle %></div>
		<%end if%>
        <div class="intro-body"><%=Ncontent %></div>
    </div>
</div>
<%
    ELSEIF rsif.RecordCount  > 1 THEN 
     CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
<div class=" container-fluid">
    <div class="container pdb-40">
        <hr class="Hr-Title">
        <div class="form-group" style="position: relative; height: 36px;">
            <div style="float: left; background: #75bb1f;" class="in-title">
                <h3 class="H-Title"><%=CName %></h3>

            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i>&nbsp; <%=CName %> </i></div>
        </div>
        <%  

    do while not rsif.eof 
        Title   = Trim(rsif("Title"))
        desc    = Trim(rsif("Author"))
        FTitle    = Trim(rsif("Description")) 
        Ncontent    = Trim(rsif("Body")) 
        cateID  = Trim(rsif("CategoryID"))
        NewsID  = Trim(rsif("NewsID"))
        nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsif("PictureId")&"'")    
        linkuri = func_uri(cateID,NewsID,Title)              
        CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
        %>
        <div class="col-md-6">
            <table style="width: 100%;">
                <tbody>
                    <tr>
                        <td style="width: 100px;">
                        <a href="<%=linkuri %>">
                            <img src="/images_upload/<%=nimg %>" class="cat-img" alt=<%=Title%> /></td>
                        </a>
                        <td style="position: relative">
                            <a href="<%=linkuri %>">
                                <h3 class="cat-title"><%=LCase(Title) %></h3>
                            </a>
                            <a href="<%=linkuri %>" class="cat-dls">Xem chi tiết..</a></td>
                    </tr>
                </tbody>
            </table>
            <hr class="cat-hr" />
        </div>
        <%     
    rsif.MoveNext
    Loop

        %>
    </div>
</div>
<%
     END IF ''IF 2
     END IF ''IF 1
End Sub
%>


<%'------------------------------------------------------------------------------------------------------------------------------------ %>
<%
Sub Fs_information(NewsID) 
    sqlN = "SELECT * FROM V_News WHERE   status = '4' And  NewsID = '"&NewsID&"'  And (url_video IS NULL or url_video ='')"
    set rsN = Server.CreateObject("ADODB.Recordset")
    rsN.open sqlN,con,1
    IF NOT rsN.eof THEN
        Title       = Trim(rsN("Title"))
        desc        = Trim(rsN("Author"))
        FTitle      = Trim(rsN("Description")) 
        Ncontent    = Trim(rsN("Body")) 
        cateID      = Trim(rsN("CategoryID"))
        NewsID      = Trim(rsN("NewsID"))
        'nimg        = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsN("PictureId")&"'")      
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&cateID&"'")
%>

<div class=" container-fluid">
    <div class="container">
        <hr class="Hr-Title">
        <div class="form-group" style="position: relative; height: 36px;">
            <div style="float: left; background: #75bb1f;" class="in-title">
                <h3 class="H-Title"><%=CName %></h3>

            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i>&nbsp; <%=CName %> </i></div>
        </div>
        <h1 class="intro-title text-center"><%=Title %></h1>
		<%if FTitle<>"" then%>
        <div class="intro-phu text-center"><%=FTitle %></div>
		<%end if%>
        <div class="intro-body"><%=Ncontent %></div>
    </div>
</div>
<%
            END IF 'not eof
            rsN.Close
%>
<%End Sub%>
<%'------------------------------------------------------------------------------------------------------------------------------------ %>


<%sub Fs_services(cid,LgID) 
 
    
%>
<div class="container-fluid">
    <%      
       sqlCate=" SELECT*  FROM  NewsCategory WHERE   LanguageId = '"&lang&"' and CategoryID = '"&cid&"' ORDER BY CategoryOrder"
       set rsCate = Server.CreateObject("ADODB.RECORDSET")
       rsCate.Open sqlCate,con,1

        CateNote =  Trim(rsCate("CategoryNote"))

    %>
    <div class="container">
        <p class="sv2-note">
            <%=CateNote %>
        </p>
    </div>
    <%
       do while not rsCate.eof
       CategoryID   =   rsCate("CategoryId")
       
       if Trim(rsCate("CategoryLink"))<>"" then 
        varLink=rsCate("CategoryLink")
       else 
            varLink= "/"&Replace(Uni2NONE(rsCate("CategoryName"))," ","")&"/"&rsCate("CategoryId")&".html"
       End if 
       query = trim(Replace(Request.QueryString("cateId")," ","+"))
       if query <> "" then
               if trim(rsCate("CategoryId")) = query then active = "active" else active = ""
       else
               if dem = 0 then active = "active" else active = "" 
               dem = dem +1
       end if
           YoungestChildren = rsCate("YoungestChildren")
           CategoryID = rsCate("CategoryID")
           CategoryName = rsCate("CategoryName")                                                
           call Fs_cate(YoungestChildren,CategoryID,varLink,CategoryName,active)                          
       rsCate.MoveNext()
       loop
    %>
</div>
<%end sub %>





<% sub Fs_cate(YoungestChildren,CategoryID,varLink,CategoryName,active)
    CID = trim(Replace(Request.QueryString("cid")," ","+"))

    IDC = Trim(Replace(CID,".html",""))
    IF IDC <> "" And IsNumeric(Trim(Replace(CID,".html",""))) THEN
            if trim(CategoryID) = IDC then 
                active = "active" 
            elseif Cint(getColVal("newscategory","ParentCategoryID","categoryid = '"&IDC&"'")) =CInt(CategoryID) then 
                active = "active" 
            else 
                active = ""
            end if
   END IF
   ' Response.Write getColVal("newscategory","ParentCategoryID","categoryid = '"&query&"'")
   ' Response.Write (YoungestChildren &" "& CategoryID )
    if YoungestChildren >0 and CategoryID <> 0 then
%>

<%
            sqlSM= "SELECT  * " &_
                   "FROM      NewsCategory " &_
                   "WHERE     (ParentCategoryID = '"&CategoryID&"') and (CategoryStatus = 2 or CategoryStatus = 4 ) And  CategoryStatus <> 1  ORDER BY CategoryOrder"
          
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1




%>
<!-- show cate  Cha-->






<%
            IF  not rsSM.eof THEN
%>

<div>
    <%
            do while not rsSM.eof

   
            call Fs_cate(rsSM("YoungestChildren"),rsSM("CategoryID"),getLink(rsSM("CategoryID"),"",rsSM("CategoryName")),rsSM("CategoryName"),"")
         
            rsSM.MoveNext()
            loop
    %>
</div>
<% 
            END IF
%>


<%
    else
%>


<div class="container pdb-20">
    <hr class="Hr-Title">
    <div style="float: left; background: #75bb1f;" class="in-title">
        <h3 class="H-Title"><%=CategoryName %></h3>
    </div>

</div>
<%
        Call Fs_CateService1(CategoryID,LgID)
%>





<%
    end if
    
%>


<% end sub %>


<%
    sub Fs_CateService1(cid_,LgID)
    sql1 = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')  Order By   LastEditedDate DESC "  
    set rsc = server.CreateObject("ADODB.RECORDSET")
    rsc.open sql1,con,1 
    Response.Write "<br />"   
    if not rsc.eof then   
        tem_ = 0
        
       
%>

<div class="container pdb-20">
    <div class="row sv-line">
        <%
            do while Not rsc.EOF
            Title   = Trim(rsc("Title"))
            desc    = Trim(rsc("Author"))

            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if


            desc_short    = Trim(rsc("DecsBannerImage"))
                
            cateID  = Trim(rsc("CategoryID"))
            NewsID  = Trim(rsc("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsc("PictureId")&"'")   
            url_img = "style= 'background:url(/images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)         
            CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")

            CateLoai= getColVal("NewsCategory","CategoryLoai","CategoryID = '"&cid_&"'") 

        %>


        <%  
            IF CateLoai  = "2" THEN
        %>
        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                    <p class="sv-title"><%=Title %></p>
                </a>
            </div>
            <p class="sv-desc"><%=desc_ %></p>
        </div>
        <%  
            ELSEIF CateLoai  = "3" THEN
        %>
        <div class="col-md-6 sv-item">
            <div class="sv-items" <%=url_img %>>
                <div style="display: table; overflow: hidden;" class="sv2-w">
                    <div style="display: table-cell; vertical-align: middle;" class="text-center ">
                        <a href="<%=linkuri %>" class="sv2-title"><%=Title %></a>
                    </div>
                </div>
            </div>
        </div>
        <%

            END IF
        %>

        <%
            rsc.MoveNext
            Loop
        %>
    </div>
</div>
<%
    set rsc = nothing
    end if 'end if dich vu 
    End sub
%>
<%'------------------------------------------------------------------------------------------------------------------------------------------------------- %>
<%
    sub Fs_CateService(cid_,IDF)
    sql1 = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryID = '"&cid_&"'  And (url_video IS NULL or url_video ='')  Order By   LastEditedDate DESC "  
    set rsc = server.CreateObject("ADODB.RECORDSET")
    rsc.open sql1,con,1 
    Response.Write "<br />"   
    if not rsc.eof then   
       CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
       CNote =  getColVal("newscategory","categoryNote","categoryid = '"&cid_&"'")               
%>
<div class="container">
    <hr class="Hr-Title">
    <div style="float: left; background: #75bb1f;" class="in-title">
        <h3 class="H-Title"><%=CName %></h3>
    </div>
    &nbsp;
    <br />
    <p class="sv-cnote clear"><%=CNote %></p>
</div>
<div class="container pdb-80">
    <div class="row sv-line">
        <%IF IDF = 2 THEN 'Giao dien 1 %>
        <p class="txt-nomal text-center">
            DANH SÁCH HÓA CHẤT SỬ DỤNG:<br>
        </p>
        <%END IF ' END GIAO DIEN %>
        <%
            do while Not rsc.EOF
            Title   = Trim(rsc("Title"))
            desc    = Trim(rsc("Author"))
            Description    = Trim(rsc("Description"))

            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if
            desc_short    = Trim(rsc("DecsBannerImage"))              
            cateID  = Trim(rsc("CategoryID"))
            NewsID  = Trim(rsc("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsc("PictureId")&"'")   
            url_img = "style= 'background:url(/images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)         
            CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")

        %>

        <%IF IDF = 1 THEN 'Giao dien 1 %>
        <div class="col-md-6 sv-item">
            <div class="sv-items" <%=url_img %>>
                <div style="display: table; overflow: hidden;" class="sv2-w">
                    <div style="display: table-cell; vertical-align: middle;" class="text-center ">
                        <a href="<%=linkuri %>" class="sv2-title"><%=Title %></a>
                    </div>
                </div>
            </div>
        </div>
        <%ELSEIF IDF = 2 THEN 'Giao dien 2 %>
        <div class="col-md-12 sv-item">
            <div class="col-md-3">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" alt="<%=Title %>" />
                </a>
            </div>
            <div class="col-md-9">
                <ul class="news-inver">
                    <li><span class="sv-cnotes">Thông tin sản phẩm:</span>
                        <p class="sv-cnotes">
                            <%=Title %> - <%=Description %>
                            <br />
                            - Can : 5L

                            <br />
                            <%=desc %>
                            <br />




                        </p>
                    </li>
                    <%
                                if desc_short <> "" or Not IsNull(desc_short) then
                    %>
                    <li><span class="sv-cnotes">Cách dùng:</span>
                        <p class="sv-cnotes"><%=desc_short %></p>
                    </li>
                    <%
                                    
                                end  if
                    %>
                </ul>

            </div>
            <hr class="sv-rowline" />
        </div>
        <%END IF ' END GIAO DIEN %>
        <%
            rsc.MoveNext
            Loop
        %>
    </div>
</div>
<%
    set rsc = nothing
    end if 'end if dich vu 
    End sub
%>
<%'------------------------------------------------------------------------------------------------------------------------------------------------------- %>





<%
    sub Fs_ServiceByLoai(idLoai,IDF)
    sqlCat = "SELECT  * FROM  NewsCategory WHERE  CategoryLoai = '"&idLoai&"' And  Ck = '1' order by CategoryID ASC "  
    set RsCat = server.CreateObject("ADODB.RECORDSET")
    RsCat.open sqlCat,con,1   
    IF NOT RsCat.EOF THEN   'IF 1  
        Do  while NOT RsCat.EOF
            Toppic = Trim(RsCat("CategoryName"))
            sqlN = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryLoai = '"&idLoai&"'  And (url_video IS NULL or url_video ='')  Order By   LastEditedDate DESC "  
            set RsN = server.CreateObject("ADODB.RECORDSET")
            RsN.open sqlN,con,1 
            IF NOT RsN.EOF THEN  
%>
<div class="container pdb-20">
    <hr class="Hr-Title">
    <div style="float: left; background: #75bb1f;" class="in-title">
        <h3 class="H-Title"><%=Toppic %></h3>
    </div>
</div>
<div class="container pdb-20">
    <div class="row sv-line">
        <%
           DO WHILE Not RsN.EOF
            Title   = Trim(RsN("Title"))
            desc    = Trim(RsN("Author"))
            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if
            desc_short    = Trim(RsN("DecsBannerImage"))               
            cateID  = Trim(RsN("CategoryID"))
            NewsID  = Trim(RsN("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&RsN("PictureId")&"'")   
            url_img = "style= 'background:url(/images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)         
            CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
            CateLoai= getColVal("NewsCategory","CategoryLoai","CategoryID = '"&cid_&"'") 

        %>

        <%IF IDF = 1 THEN %>



        <div class="col-md-6 sv-item">
            <div class="sv-items" <%=url_img %>>
                <div style="display: table; overflow: hidden;" class="sv2-w">
                    <div style="display: table-cell; vertical-align: middle;" class="text-center ">
                        <a href="<%=linkuri %>" class="sv2-title"><%=Title %></a>
                    </div>
                </div>
            </div>
        </div>

        <%ELSEIF IDF = 2 THEN %>
        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                    <p class="sv-title"><%=Title %></p>
                </a>
            </div>
            <p class="sv-desc"><%=desc_ %></p>
        </div>
        <%END IF %>

        <%
            RsN.MoveNext
            Loop
        %>
    </div>
</div>
<%
   set RsN = nothing
   END IF  'End IF 2
%>
<%                
    RsCat.MoveNext
    Loop
    set RsCat = nothing
    END IF 'End IF 1 
End sub
%>



<%
    sub Fs_ServiceCommon(idLoai,IDF)  
            sqlN = "SELECT  * FROM V_News WHERE   status = '4' And  CategoryLoai = '"&idLoai&"'  And (url_video IS NULL or url_video ='')  Order By   LastEditedDate DESC "  
            set RsN = server.CreateObject("ADODB.RECORDSET")
            RsN.open sqlN,con,1 
            IF NOT RsN.EOF THEN  
%>
<div class="container pdb-20">
    <hr class="Hr-Title">
    <div style="float: left; background: #75bb1f;" class="in-title">
        <h3 class="H-Title">Dịch vụ cung cấp</h3>
    </div>
</div>
<div class="container pdb-20">
    <div class="row sv-line">
        <%
           DO WHILE Not RsN.EOF
            Title   = Trim(RsN("Title"))
            desc    = Trim(RsN("Author"))
            if  Len(desc)  < 95  then 
                desc_ = desc
            else
                desc_ = Left(desc,95)&"..."
            end if
            desc_short    = Trim(RsN("DecsBannerImage"))               
            cateID  = Trim(RsN("CategoryID"))
            NewsID  = Trim(RsN("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&RsN("PictureId")&"'")   
            url_img = "style= 'background:url(/images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)         
            CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'")
            CateLoai= getColVal("NewsCategory","CategoryLoai","CategoryID = '"&cid_&"'") 

        %>

        <%IF IDF = 1 THEN %>



        <div class="col-md-6 sv-item">
            <div class="sv-items" <%=url_img %>>
                <div style="display: table; overflow: hidden;" class="sv2-w">
                    <div style="display: table-cell; vertical-align: middle;" class="text-center ">
                        <a href="<%=linkuri %>" class="sv2-title"><%=Title %></a>
                    </div>
                </div>
            </div>
        </div>

        <%ELSEIF IDF = 2 THEN %>
        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                    <p class="sv-title"><%=Title %></p>
                </a>
            </div>
            <p class="sv-desc"><%=desc_ %></p>
        </div>
        <%END IF %>

        <%
            RsN.MoveNext
            Loop
        %>
    </div>
</div>
<%
   set RsN = nothing
   END IF  'End IF 2
%>
<%                
End sub
%>


<%
    sub Fs_ServiceDetail(IFID,NewsID)  
    IF NewsID <> "" AND  IsNumeric(NewsID) THEN    
        sqlN = "SELECT * FROM V_News WHERE   status = '4' And  NewsID = '"&NewsID&"'  And (url_video IS NULL or url_video ='')"
        set rsN = Server.CreateObject("ADODB.Recordset")
        rsN.open sqlN,con,1
        IF NOT rsN.eof THEN
            Title       = Trim(rsN("Title"))
            desc        = Trim(rsN("Author"))
            FTitle      = Trim(rsN("Description")) 
            Ncontent    = Trim(rsN("Body")) 
            cateID      = Trim(rsN("CategoryID"))
            NewsID      = Trim(rsN("NewsID"))
            'nimg        = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsN("PictureId")&"'")      
            CName       =  getColVal("newscategory","categoryname","categoryid = '"&cateID&"'")
%>
<div class=" container-fluid">
    <div class="container">
        <%IF  IFID = 1 THEN %>
        <hr class="Hr-Title">
        <div style="float: left; background: #75bb1f;" class="in-title">
            <h1 class="H-Title"><%=Title %></h1>
        </div>
        <div class="intro-body " style="clear: both;"><%=Ncontent %></div>
        <%ELSEIF  IFID = 2 THEN %>


        <%END IF %>
    </div>
</div>
<%
            END IF 'not eof       
            rsN.Close
    END IF ' IF 1
%>
<%end  sub %>
<%'------------------------------------------------------------------------------------------------------------------------------------ %>


<% sub  Fs_SevCungLoai(cid_,Nid)
    sql = "SELECT  * FROM V_News WHERE    status = '4' And   CategoryID = '"&cid_&"' And NewsID <> '"&Nid&"'  And (url_video IS NULL or url_video ='')    Order By   LastEditedDate DESC "  
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1
    
    Response.Write "<br />"   
    if not rsi.eof then   
        tem_ = 0
        Title   = Trim(rsi("Title"))
        desc    = Trim(rsi("Author"))

        if  Len(desc)  < 95  then 
            desc_ = desc
        else
            desc_ = Left(desc,95)&"..."
        end if
        desc_short    = Trim(rsi("DecsBannerImage"))        
        cateID  = Trim(rsi("CategoryID"))
        NewsID  = Trim(rsi("NewsID"))
        nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
        linkuri = func_uri(cateID,NewsID,Title)         
%>
<div class="container">
    <hr class="Hr-Title" />
    <div class="form-group" style="position: relative; height: 36px;">
        <div style="float: left; background: #75bb1f;" class="in-title">
            <h3 class="H-Title">Dịch vụ cùng chuyên mục</h3>
        </div>
    </div>

    <div class="row sv-line">
        <%
            do while Not rsi.EOF
        %>
        <div class="col-md-3 sv-item">
            <div class="sv-items">
                <a href="<%=linkuri %>">
                    <img class="img-responsive" src="/images_upload/<%=nimg %>" />
                    <p class="sv-title"><%=Title %></p>
                </a>
            </div>
            <p class="sv-desc"><%=desc_ %></p>
        </div>
        <%
            rsi.MoveNext
            Loop
        %>
    </div>

</div>
<%
    set rsi = nothing
    end if 'end if dich vu
    end sub
%>

<%'------------------------------------------------------------------------------------------------------------------------------------ %>
<%
    sub Fs_Patner() 
    
    sqlp = "SELECT  * FROM patner WHERE  [view] = 1 Order By [DateCreate] "
    Set RsP=Server.CreateObject("ADODB.Recordset")
    RsP.open sqlp,con,1
    IF NOT RsP.EOF THEN
      
%>
<div class="container">

    <hr class="Hr-Title" />
    <div class="form-group" style="position: relative; height: 36px;">
        <div style="float: left; background: #75bb1f;" class="in-title">
            <h3 class="H-Title">Khách Hàng tiêu biểu</h3>
        </div>
    </div>
    <div class="Patner">
        <%
            Do while  NOT RsP.EOF

               imgv = Trim(Rsp("AvImg"))  
                
               if imgv <> ""  then
                    
               
        %>
        <div style="padding: 2px;">
            <div class="pnl-item">
                <img src="/images_upload/IMG_Customer/<%=imgv %>" class="img-responsive " />
            </div>
        </div>
        <%end  if
            RsP.MoveNext
            Loop
        %>
    </div>

</div>
<%END IF %>

<%end sub %>








<%
    sub  Ykien()
        sqlyk = "SELECT * FROM  Y_KIEN  WHERE  faq = 3  AND   show = '1' AND    NewsId IS Not Null"   
        set rsFaq = Server.CreateObject("ADODB.Recordset")
        rsFaq.open sqlyk,con,1
        IF NOT rsFaq.EOF THEN
           
       


%>


<div class="container pt-40">
    <h2 class="in-title">Khách Hàng Nói Gì Về Chúng Tôi</h2>
    <p class="br-line"></p>
    <br>
    &nbsp;
    <div class="product-other">
    </div>
</div>
<div class="" style="position: relative;">
    <div id="Ykien_comment" class="carousel slide" data-ride="carousel">
        <!-- Indicators -->
        <ol class="carousel-indicators">
            <li data-target="#Ykien_comment" data-slide-to="0" class="active"></li>
            <li data-target="#Ykien_comment" data-slide-to="1"></li>
            <li data-target="#Ykien_comment" data-slide-to="2"></li>
        </ol>
        <div class="container">
            <!-- Wrapper for slides -->
            <div class="carousel-inner " role="listbox" style="">
                <%
                    stt = 0
                    DO WHILE NOT rsFaq.EOF   
                        Fname = Trim(rsFaq("hovaten"))
                        TrLoi = Trim(rsFaq("Traloi"))
                        avata    = Trim(rsFaq("Photo"))
                        NID    = Trim(rsFaq("NewsID"))
                        if stt = 0 then 
                            ac = "active"
                        else
                            ac = ""
                        end  if

                    IF (NID <> "" or Not IsNull(NID)) AND  IsNumeric(NID) THEN

                     sqlN = "SELECT    LargePictureFileName   FROM Picture  WHERE PictureID = (SELECT PictureID FROM News WHERE NewsID ='"&NID&"' ) "   
                     set rsFN = Server.CreateObject("ADODB.Recordset")
                     rsFN.open sqlN,con,1
                     IF NOT rsFN.EOF THEN
                        IMG =    Trim(rsFN("LargePictureFileName"))
                       
                     ELSE
                        IMG ="ads-dls.png"
                     END IF
                     ProName_ = getColVal("V_News","title","NewsID = '"&NID&"'") 
                    END IF


                %>
                <div class="item <%=ac %>  ">
                    <img src="/images_upload/<%=IMG %>" class="img-responsive" />
                    <div class="carousel-caption Title-Yk"><%=ProName_ %></div>
                    <br />
                    <table class="yk-txt" style="margin: auto auto" border="0">
                        <tr>
                            <td>
                                <p class="yk-txt">
                                    Tên khách hàng: <br />
                                    <%=Fname %><br />
                                    Ngân sách: 500.000 VNĐ<br />
                                </p>
                            </td>
                            <td>
                                <p class="yk-avata">
                                    <img src="/images_upload/IMG_Customer/<%=avata %>" class="yk-img" />
                                </p>
                            </td>
                            <td>
                                <img src="/images/icon/rs-right.png" style="height: 170px; margin: auto 10px;" />
                            </td>
                            <td style="width: 400px;">
                                <p class="yk-txt">
                                    <%=TrLoi %>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <br />
                </div>
                <%
                    stt = stt + 1
                    rsFaq.MoveNext
                    Loop 
                %>
            </div>
        </div>

    </div>
    <!-- Controls -->
    <a class="left carousel-control" href="#Ykien_comment" role="button" data-slide="prev">
        <span class="glyphicon glyphicon-chevron-left" aria-hidden="true"></span>
        <span class="sr-only">Previous</span>
    </a>
    <a class="right carousel-control" href="#Ykien_comment" role="button" data-slide="next">
        <span class="glyphicon glyphicon-chevron-right" aria-hidden="true"></span>
        <span class="sr-only">Next</span>
    </a>
</div>
    

<%
    END IF
    end sub 
%>

<%Sub team() %>
<!-- Team Container -->
<div class="w3-container w3-padding-64 w3-center" id="team">
<h2>OUR TEAM</h2>
<p>Meet the team - our office rats:</p>

<div class="w3-row"><br>



<div class="w3-quarter">
  <img src="/images/ThuyPTT.jpg" alt="Boss" style="width:45%" class="w3-circle w3-hover-opacity">
  <h3><a href="/">PGS. TS Phạm Thị Thanh Thùy</a></h3>
  <p>Founder & owner</p>
</div>


<div class="w3-quarter">
  	<img src="/images/tuannv.jpg" alt="Boss" style="width:45%" class="w3-circle w3-hover-opacity">
    <h3><a href="http://jannguyen.com/" target="_blank">CTO. Nguyễn Tuân</a></h3>
  <p>Developer</p>
</div>

<div class="w3-quarter">	
  <img src="/images/AnLN.jpg" alt="Boss" style="width:45%" class="w3-circle w3-hover-opacity">
  <h3><a href="https://www.daihockhonggiay.com/" target="_blank">ThS. Lê Ngọc An </a></h3>
  <p>Senior advisor</p>
</div>

<div class="w3-quarter">
  <img src="/images/ThanhTV.jpg" alt="Boss" style="width:45%" class="w3-circle w3-hover-opacity">
  <h3>ThS. Trần Vĩnh Thanh</h3>
  <p>Senior advisor</p>
</div>

</div>
</div>
        
        
<%End Sub %>

<%Sub Fs_Footer() %>
      <footer class="footer">
    <div class="container">
        <div class="row">
            <div class="footer-col col-lg-3 col-sm-12">
                    <h4><%=TitleF %> </h4>
                    <ul>
                        <li style="text-align: justify"><% =introduction %></li>
                    </ul>
                </div>
        <div class="footer-col col-lg-3 w3-hide-small ">
                <% call Hotro() %>
                </div>
             <div class="footer-col col-lg-3 w3-hide-small">
                        <iframe src="https://www.facebook.com/plugins/page.php?href=https%3A%2F%2Fwww.facebook.com%2FToilathaomocVN%2F&tabs=timeline&width=340&height=500&small_header=false&adapt_container_width=true&hide_cover=false&show_facepile=true&appId"
                        style="border:none;overflow:hidden; aspect-ratio: 1 / 1; max-width: 100%" scrolling="no" frameborder="0" allowfullscreen="true" allow="autoplay; clipboard-write; encrypted-media; picture-in-picture; web-share"></iframe>
                </div>
              <div class="footer-col col-lg-3 col-sm-12">
                    <h4>follow us</h4>
                    <div class="social-links">
                        <a href="<%=idfacebook %>"><i class="fab fa-facebook-f"></i></a>
                        <a  href="<%=idgplus %>"><i class="fa fa-google-plus"></i></a>
                        <a href="#"><i class="fab fa-linkedin-in"></i></a>
                        <a href="<%=idyoutube %>"><i class="fa fa-youtube-play"></i></a>
      <a href="Tel:<%=Tel_sys %>"><i class="fa fa-phone-square"></i></a>
                    </div>
                  <div class="  text-left social-links-two">
                      <%if Hotline<>"" then%>
                      <a  href="Tel:<%=Hotline %>" class="w3-text-white "><i class="fa fa-phone-square"></i> Hotline: <%=Hotline %></a>
                      <%end if%>
                      <a  class="w3-text-white " href="mailto:<%=Email %>"><i class="fa fa fa-envelope-o"></i> <%=Email %></a> <br>
                  </div>
                </div>
        </div>
</footer>
<%End Sub%>
<% sub Hotro()%>
      <h4>
      <% 
          sqlMenu="SELECT    * " &_
                             "FROM      NewsCategory " &_
                             "WHERE   LanguageId = '"&lang&"' and  (ParentCategoryID = 0) and (CategoryStatus = 2) AND  CategoryHome='8'  ORDER BY CategoryOrder"
          set rsMenu = Server.CreateObject("ADODB.RECORDSET")
                     rsMenu.Open sqlMenu,con,1
               	do while not rsMenu.eof
                    CategoryID  =   rsMenu("CategoryId")
                    if Trim(rsMenu("CategoryLink"))<>"" then 
                        varLink=rsMenu("CategoryLink")
                    else 

                        varLink= "/"&Replace(Uni2NONE(rsMenu("CategoryName"))," ","")&"/"&rsMenu("CategoryId")&".html"

                    End if 
                    query = trim(Replace(Request.QueryString("cateId")," ","+"))
					YoungestChildren = rsMenu("YoungestChildren")
					CategoryName = rsMenu("CategoryName")                                                 
				    rsMenu.MoveNext()
                    response.Write(CategoryName)
				loop%>
          
    </h4>
      <ul>
        <%Call conHotro(YoungestChildren,CategoryID,CategoryName) %>  
        </ul>
<%end sub %>

  <% sub conHotro(YoungestChildren,CategoryID,CategoryName) %>
     
<% 
   sqlSM= "SELECT  * " &_
                   "FROM      NewsCategory " &_
                   "WHERE     (ParentCategoryID = '"&CategoryID&"')  And  CategoryStatus <> 1  ORDER BY CategoryOrder"    
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1  %>
            
           <%	do while not rsSM.eof
                    CategoryName = rsSM("CategoryName")
               if Trim(rsSM("CategoryLink"))<>"" then 
                        varLink=rsSM("CategoryLink")
                    else 

                        'varLink= "/"&Replace(Uni2NONE(rsSM("CategoryName"))," ","-")&"/"&rsSM("CategoryId")&".html"
                        varLink = getLink(rsSM("CategoryID"),"",rsSM("CategoryName"))
                    End if 
               response.Write ( "<li>"&"<a href=" &varlink&">" & CategoryName & "</a>" & "</li>"  )
               rsSM.MoveNext()
				loop%>
      
     
       
        <% end sub %>
 

<%
''**********************************************************************
'' Function: write_Ads
'' Version: 1  
'' Date: 2019/08/14
'' Author: hoavm    
'' Description: Function show in page Ads
''case 1    Ads_Position="Slider ảnh sau menu(1600x400)"    
''case 2    Ads_Position="Text trong bài viết"
''case 3    Ads_Position="Video nổi bật(1280x720)"
''case 4    Ads_Position="Video dưới nổi bật(1280x720)" 
''**********************************************************************
Sub write_Ads1(CatId,lang,AdsPosition,AdsCount,AdsColumn)
    'AdsDirection=0=Chiều dọc
    '            =1=Chiều ngang
    Dim rsAds
    set rsAds=Server.CreateObject("ADODB.Recordset")
    if AdsCount>0 then
        sqlAds="SELECT  top " & AdsCount
    else
        sqlAds="SELECT"
    end if
    sqlAds=sqlAds & " Ads_id, Ads_Title, Ads_Link, Ads_ImagesPath, " &_
            "       Ads_Type, Ads_width, Ads_height, Ads_Target,Ads_Note " &_
            "FROM   V_Ads " &_
            "WHERE  (Ads_LangID ='"& lang &"') and (Ads_Position = " & AdsPosition & ") and ( " &_
            "       (CategoryId=0) or (CategoryId=" & CatId & ")"
    strCat=Trim(GetListParentCat(CatId))
    if strCat<>"" then
        ArrCat=Split(" " & strCat & " ")
        for i=1 to UBound(ArrCat)
            if IsNumeric(ArrCat(i)) then
                sqlAds=sqlAds & " or (CategoryId=" & ArrCat(i) & " and Ads_OnlineChildren=1)"
            end if
        next
    end if
    sqlAds=sqlAds & ") "
    response.write sqlAds
    rsAds.open sqlAds,con,3
    if rsAds.eof then
        rsAds.close
        set rsAds=nothing
        Exit sub
    end if
        
    if clng(AdsPosition)=3 then
        ytube_class="ytube_lates"
    else
        if clng(AdsPosition)=4 then
            ytube_class="ytube_prior"
        end if
    end if  
          
    'Dùng vòng lặp check kiểu Ads_Type thiết lập là loại nào
    i=1
    Do while not rsAds.eof
        'Ads_Type:  0: GIF,JPG
        '           1: Youtube
        '           2: txt
        Ads_id          =   clng(rsAds("Ads_id"))
        Ads_Title       =   trim(rsAds("Ads_Title"))
        Ads_width       =   rsAds("Ads_width")
        Ads_height      =   rsAds("Ads_height")
        Ads_Link        =   Trim(rsAds("Ads_Link"))
        itarget         =   Trim(rsAds("Ads_Target"))            
        Ads_ImagesPath  =   rsAds("Ads_ImagesPath")
        Ads_ImagesPath  =   NewsImagePath&Ads_ImagesPath
        Ads_Note        =   Trim(rsAds("Ads_Note"))

    '====================================================================
    Select Case Clng(rsAds("Ads_Type"))
        Case 0 'GIF,JPG show slide
            if i=1 then
                Response.write("<section class='home-slider owl-carousel'>")
            end if
            Response.write "<div class='slider-item' style='background-image:url("&Ads_ImagesPath&");'>"      
            Response.write "<div class='overlay'></div>"           
            Response.write "<div class='container'>"           
            Response.write "<div class='row no-gutters slider-text align-items-center justify-content-start' data-scrollax-parent='true'>"           
            Response.write "<div class='col-md-6 ftco-animate'>"               
            Response.write "</div>"           
            Response.write "</div>"           
            Response.write "</div>"           
            Response.write "</div><!---/.slider-item--->" 
            if i=rsAds.recordcount  then
                Response.write "</section>"
                Response.write "<style type='text/css'>.owl-carousel.home-slider .slider-item {height: "&Ads_height&"px;}</style>"
            end if    
        '=========================================================================================================                  
        Case 1 'Youtube,vimeo           
            if i=1 then 
                       
        %>
             <div class='porto-wrap-container container'>
             <div class='row'>
             <div class='vc_column_container col-md-12 d-xl-block d-none vc_custom_1566027133036 mt-0 mb-0 pt-0 pb-0 section-no-borders'>
             <div class='wpb_wrapper vc_column-inner'>
             <div class='wpb_raw_code wpb_content_element wpb_raw_html'><div class='wpb_wrapper'></div></div>
             <div class='<%=ytube_class%> porto-carousel owl-carousel has-ccols ccols-xl-4 ccols-lg-4 ccols-md-3 ccols-sm-2 ccols-1 doi_tac owl-loaded owl-drag'>           
        <%
            '=============================================
            end if
        %>           
            <div class='wpb_single_image wpb_content_element vc_align_center'>
                <div class='wpb_wrapper'>
                <a href='<%=Ads_Link %>' class='popup-vimeo'>
                    <div class='vc_single_image-wrapper vc_box_border_grey'>
                        <img width='<%=Ads_width%>' height='<%=Ads_height%>' src='<%=Ads_ImagesPath%>' alt='<%=Ads_Title%>' title='<%=Ads_Title%>'/>                                  
                        <%if AdsPosition=3 then %>
                        <span class='ytube_play'></span>
                        <%end if %>
                    </div>
                </a>
                </div>
            </div>
        <%if i=rsAds.recordcount  then%>
                    </div><!---/.ytube_later-->                
                </div>
            </div>          
            </div>
            </div><!---/.porto-wrap-container--->          
        <%end if 
        '===========================================================================================================
        Case 2 'text
        if i=1 then
            Response.write "<section class='ftco-services ftco-no-pb'>"
            Response.write "<div class='container-wrap'>"
                Response.write "<div class='row no-gutters'>"
        end if       
            'show item loop
            Response.write "<div class='col-md-3 d-flex services align-self-stretch py-5 px-4 ftco-animate "&Ads_background&"'>"
                Response.write "<div class='media block-6 d-block text-center'>"
                    Response.write "<div class='icon d-flex justify-content-center align-items-center'>"                        
                    Response.write "</div>"
                    Response.write "<div class='media-body p-2 mt-3'>"
                        Response.write "<h3 class='heading'>"&Ads_Title&"</h3>"
                        Response.write "<p>"&Ads_Note&"</p>"
                    Response.write "</div>"
                Response.write "</div><!---/.media block-6--->"      
            Response.write "</div><!---/.col-md-3--->"
        if i=rsAds.recordcount  then
               Response.write "</div><!---/.row--->"
           Response.write "</div><!---/.container-wrap--->"
           Response.write "</section><!--/.section--->"
        end if 
    End Select 'end select Ads_Type
    i=i+1
    rsAds.movenext
    Loop    
    rsAds.close
    set rsAds=nothing  
End Sub
%>

<%

function G_CategoryLoai(Loai)
    sql_s="SELECT CategoryID FROM NewsCategory WHERE CategoryLoai='"&Loai&"'"
    'sql_s ="SELECT * FROM NewsCategory"
    Set rs_s=Server.CreateObject("ADODB.Recordset")
    rs_s.open sql_s,con,1
    if not rs_s.EOF then
        G_CategoryLoai  =   rs_s("CategoryID")
    end if
    set rs_s = nothing
end function  %>
  <!-- backtoTop -->
   <%Sub backTop() %>
      <button id="topBtn">
          <i class="fas fa-arrow-up"></i>
      </button>
      <script type="text/javascript">
          $(document).ready(function () {

              $(window).scroll(function () {
                  if ($(this).scrollTop() > 40) {
                      $('#topBtn').fadeIn();
                  } else {
                      $('#topBtn').fadeOut();
                  }
              });

              $("#topBtn").click(function () {
                  $('html ,body').animate({ scrollTop: 0 }, 800);
              });
          });
          </script>
      <%End Sub %>
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
         

