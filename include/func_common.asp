<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/include/func_DateTime.asp" -->
<!--#include virtual="/include/function_toan.asp"-->
<%
    lang = Session("Language")
    IF lang = "" THEN lang = "VN"   
%>

<%Sub menu()%>
<table>
    <tr>
        <td>

            <ul>
                <%           
                    sqlMenu="SELECT * FROM      NewsCategory WHERE     (ParentCategoryID = 0) and (CategoryStatus = 2 or CategoryStatus = 4 )  ORDER BY CategoryOrder"
                    set rsMenu = Server.CreateObject("ADODB.RECORDSET")
                    rsMenu.Open sqlMenu,con,1
                    do while not rsMenu.eof
                    CategoryID  =   rsMenu("CategoryId")
    
                    if Trim(rsMenu("CategoryLink"))<>"" then 
                        varLink=rsMenu("CategoryLink")
                    else 
                        varLink= "/"&rsMenu("CategoryId")&"/"&Replace(Uni2NONE(rsMenu("CategoryName"))," ","-")
                    End if 
                %>
                <li><a href="<%=varLink %>"><%=rsMenu("CategoryName") %></a></li>
                <%
                            rsMenu.MoveNext()
                            loop
                %>
            </ul>


        </td>
    </tr>
</table>
<%end sub %>


<%
sub getSearch(Keyword)

  %>
     <div class="col-xs-12 col-sm-12 col-md-7 col-lg-7 w3-padding-right">
  <%
         
    Keyword=          replace(Trim(Keyword),"  "," ")   
    'if Keyword = "" then
    if false then
       
    else
    'StringSearch=    replace(Trim(Keyword),"   "," ")
    StringSearch=Keyword
    StringStaus =   StringSearch
    
 '   Response.Write Keyword&chr(13)&"<br/>"
 '   Response.Write StringSearch&"<br/>"

    dim dictrs
    Set dictrs  =   Server.CreateObject("Scripting.Dictionary")
    Set dictSearch  =   Server.CreateObject("Scripting.Dictionary")
    Set titleSearch =   Server.CreateObject("Scripting.Dictionary")
    Set rs=server.CreateObject("ADODB.Recordset")
    
    sql="SELECT * FROM  V_News  where status = 4  order by CreationDate desc "

   ' StringTitle    =   UCASE(StringSearch)
    'sql= sql&" Where ({fn UCASE(Title)} like N'%"&StringTitle&"%')"
    'Response.Write sql       
    rs.Open sql,con,1
    if rs.EOF then
      '  Response.write "Không tìm thấy kết quả nào cho từ khóa """&keyword &"""<br/>"
    else
      '  Response.write "Tìm thấy kết quả cho từ khóa """&keyword &"""<br/>"
      ' Response.write "<br/>"
        do while not rs.EOF 
   ' Response.Write rs("NewsID")&" <b> "&rs("title")&" </b>"&"<br/>"
    iItemID =   rs("NewsID")
    title   =   rs("title")
    if len(Title) >= len(Keyword) then
            iPos    =   InStr(UCASE(Title),UCASE(Keyword))
            iF iPos >= 1 then
                iItemID =   rs("NewsID")
                StringSearch    =   Mid(rs("Title"),iPos,Len(Keyword))
                if dictSearch.Exists(iItemID)=false then
                    dictSearch.Add iItemID, StringSearch
                    titleSearch.Add iItemID, Title
                    dictrs.Add iItemID, title&"|"&iItemID
       ' showresult(title&"|"&iItemID,rs("categoryid"),rs("size"),rs("price"),rs("unit"),rs("direction"),rs("SizeFacade"),rs("author"),rs("phone"),rs("address"),rs("pictureid"))
                 
                end if                  
            end if
            end if
        if dictSearch.count > 501 then
            exit do
        end if  
    rs.MoveNext
        loop

  
  '  Response.Write dictSearch.count
 '-------------------------------------Không dấu---------------------------------------------------
    if dictSearch.count =0 then
    rs.MoveFirst
    do while not rs.EOF 
    iItemID         =   rs("NewsID")
    Title= rs("title")
    if dictSearch.Exists(iItemID)=false then
    if len(Title) >= len(Keyword) then
            iPos    =   InStr(Uni2NONE(UCASE(Title)),Uni2NONE(UCASE(Keyword)))
            iF iPos >= 1 then
                iItemID         =   rs("NewsID")
                StringSearch    =   Mid(rs("Title"),iPos,Len(Keyword))
                
                    dictSearch.Add iItemID, StringSearch
                    titleSearch.Add iItemID, Title
                    dictrs.Add iItemID, title&"|"&iItemID
       
                end if                  
            end if
    end if  

    rs.MoveNext
    
    if dictSearch.count > 501 then
            exit do
        end if
    loop
    end if
      '  Response.Write dictSearch.count
 '------------------------------------đổ chố ---------------------------------------------------

        if dictSearch.count =0 then
    rs.MoveFirst
    do while not rs.EOF 
   iItemID  =   rs("NewsID")
   title =  rs("title")
    if dictSearch.Exists(iItemID)=false then
    if len(rs("Title")) >= len(Keyword) then
     '  Response.Write Keyword&" -------------  "&rs("Title")&"<br/>"
                    wordPercent =   simalarString(Uni2NONE(UCASE(rs("Title"))),Uni2NONE(UCASE(Keyword)))
   ' Response.Write wordPercent&"%"&Keyword&" -------------  "&rs("Title")&"---------------"&Uni2NONE(UCASE(rs("Title")))&" ------ "&Uni2NONE(UCASE(Keyword))&"<br/>"
                iF wordPercent >= 90 then
                        
                        StringSearch    =   rs("Title")
                        
                            dictSearch.Add iItemID, Keyword
                            'SortSearch.add iItemID, wordPercent
                            titleSearch.Add iItemID,StringSearch
                            dictrs.Add iItemID, title&"|"&iItemID


                        end if                  
                    end if
                end if  
    rs.MoveNext
    
    if dictSearch.count > 501 then
            exit do
        end if
    loop
    end if
    '        Response.Write dictSearch.count
 '------------------------------------đổ chố 2 ---------------------------------------------------

  '      if dictSearch.count =0 then
  '  rs.MoveFirst
  '  do while not rs.EOF 
  '  iItemID    =   rs("NewsID")
  '  title =    rs("title")
  '  if dictSearch.Exists(iItemID)=false then
  '  if len(rs("Title")) >= len(Keyword) then
  '   '  Response.Write Keyword&" -------------  "&rs("Title")&"<br/>"
    '               wordPercent =   simalarString(Uni2NONE(UCASE(rs("Title"))),Uni2NONE(UCASE(Keyword)))
  ' ' Response.Write wordPercent&"%"&Keyword&" -------------  "&rs("Title")&"---------------"&Uni2NONE(UCASE(rs("Title")))&" ------ "&Uni2NONE(UCASE(Keyword))&"<br/>"
    '           iF wordPercent >= 80 then
    '                   
    '                   StringSearch    =   rs("Title")
    '                   
    '                       dictSearch.Add iItemID, Keyword
    '                       'SortSearch.add iItemID, wordPercent
  '                          titleSearch.Add iItemID,StringSearch
  '                  dictrs.Add iItemID, title&"|"&iItemID&"|"&rs("categoryid")&"|"&rs("size")&"|"&rs("price")&"|"&rs("unit")&"|"&rs("direction")&"|"&rs("SizeFacade")&"|"&rs("author")&"|"&rs("phone")&"|"&rs("address")&"|"&rs("pictureid")
  '
  '
    '                   end if                  
    '               end if
    '           end if  
  '  rs.MoveNext
  '  
  '  if dictSearch.count > 501 then
    '       exit do
    '   end if
  '  loop
  '  end if
  '  Response.Write dictSearch.count&"<br/>"


    
      end if

   'call ReSearchOther(dictSearch,page)

  if Round(dictSearch.count/20) > 0 then

        Set Session("dictSearch") = dictSearch
        for i = 1 to Round(dictSearch.count/20)
        %>
        <a href="/<%=CatID%>/<%=lang%>/<%=i%>/<%=keyword%>" target="_parent">
        <%
        if i = page then
          Response.Write("<font style=""BACKGROUND-COLOR: #FFFF66"">"&i&"</font>")
        else
          Response.Write(i)
        end if
        %>|</a>
        <%
        next          

  end if     




    if dictSearch.Count >0 then
   ' Response.Write "<br/><br/>List <b>"&dictSearch.Count &"</b> keyword match Keyword: """&Keyword&"""<br/>"
    Response.Write " <input type=""hidden"" id=""countSearchRS"" name=""country"" value="""&dictSearch.Count&""">"
        'Session("searchrs") = dictRS
    else
    Response.Write " <input type=""hidden"" id=""countSearchRS"" name=""country"" value="""&dictSearch.Count&""">"
    Response.Write  "Không tìm thấy kết quả nào"
    end if
    if dictSearch.Count>0 then
        dem = 0
   karr= dictSearch.Keys()
      arr = dictSearch.Items()
    tarr = titleSearch.Items()
        count = CInt(dictSearch.Count)
    
         page = 1
         pagec = 0
    stt = 1
    do while dem < dictSearch.Count
     '  Response.Write karr(dem)&" <b>---"& arr(dem)&"---<b>"&tarr(dem)&"</b></b><br/>"
     '  Response.Write dictRS.Items()(dem)&"</b></b><br/>"
        itemArr = Split(dictRS.Items()(dem),"|")
     'Response.Write UBound(itemArr)&"</b></b><br/>"   
        title = itemArr(0)
        iItemID =  itemArr(1)
       
   if countViewPerPage <> "" then
    pagecTotal = CInt(countViewPerPage)
    else
    pagecTotal = 15
    end if
   
    sql = "Select * from V_News where status = 4 and newsid = '"&iItemID&"'"
    set rs = Server.CreateObject("ADODB.RECORDSET")
    rs.open sql,con,1
    if not rs.eof then
   link =   getLink(rs("Categoryid"),rs("newsid"),rs("title"))


%>

  <span class="sl-news-Tl"><%=stt %></span>.<a class="sl-news-Tl"sz href="<%=link %>"><%=rs("title") %></a><br> <br>
    
<%



    end if

'Response.Write page
      '  if dem < 15 then
       ' call showresult(title,iItemID,catid,size,price,unit,huong,mt,author,phone,address,picid,page)

      '  end if
    stt = stt +1
    dem = dem  +1
    pagec = pagec + 1  
        if pagec = pagecTotal then
        page = page+1
        pagec=0
        end if

    loop

    %>


  </div>
  
    <div class="col-xs-12 col-sm-12 col-md-5 col-lg-5 w3-padding-left">
        <%
        call siderbar(10)
        %>
    </div>
    <%
     

    end if
    
    end if
        end sub
%>


<%Sub phan_trang() %>
 <%if Round(dictSearch.count/20) > 0 then %>

      <%
        Set Session("dictSearch") = dictSearch
        for i = 1 to Round(dictSearch.count/20)
        %>
        <a href="/<%=CatID%>/<%=lang%>/<%=i%>/<%=keyword%>" target="_parent">
        <%
        if i = page then
          Response.Write("<font style=""BACKGROUND-COLOR: #FFFF66"">"&i&"</font>")
        else
          Response.Write(i)
        end if
        %>|</a>
        <%
        next
      %>          

  <%end if%>      



<% End sub%>



<%Sub siderbar(count,cateID)

  sqlNews="SELECT top "&count&" CategoryName,IsHomeNews,CategoryId,NewsID, CreationDate, NewsID, Title, Description, PictureID  FROM V_News WHERE  (categoryID ="&cateID&") order by CreationDate desc"
    set rsNews=Server.CreateObject("ADODB.Recordset") 
  rsNews.open sqlNews,con,3
  if not rsNews.eof then
   %>
 <div class="in-title">
            <h3 class=" H-Title"><%=rsNews("CategoryName") %></h3><!--H-gioithieu--> 
 </div>
   <%

  st  = 1
  Do while not rsNews.eof
        Title       =   Trim(rsNews("Title"))                          
        nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsNews("PictureId")&"'")   
        linkuri = func_uri(rsNews("CategoryID"),rsNews("NewsID"),rsNews("Title"))
%>      
        <div class="Item-news row pt-3 ">
                <a href="<%=linkuri %>" class="col-6 pr-0"><img class="img-responsive-self " src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
                <a href="<%=linkuri %>" class="sl-news-Tl col-6"><%=Title %></a>
        </div>
        <hr>
<%    
        st  = st + 1
    rsNews.movenext  
    loop
    %>
<%

    end if

End sub%>

<%sub listSearchProduct() %>
<select onchange="alert(this.value);">
    <%
    sqlSP = "Select * from Product where CateID in (select ID from Tb_option where ParentId =59) "
    set rsSP = Server.CreateObject("ADODB.RECORDSET")
    rsSP.Open sqlSP,con,1
    do while not rsSP.eof
    %>
    <option value="<%=rsSP("ProductIDUse") %>"><%=getColVal("tb_option","OpName","id = '"&rsSP("CateID")&"'")  %></option>

    <%
    rsSP.MoveNext()
    loop
    %>
</select>
<%end sub %>

<% 
function readFile(fileName)
        Set fs=Server.CreateObject("Scripting.FileSystemObject")
        Set f=fs.GetFile(Server.MapPath(fileName))
        Set ts = f.OpenAsTextStream(1)
         if ts.AtEndOfStream then
             readFile  = ""
        else
             readFile  = ts.ReadLine           
        end if 
        ts.Close
        set f=nothing
        set fs=nothing
        set ts = nothing
     end function 
%>
<%sub Counter() %>
<div class="hot-product-header">
    <%=ms_tk %>
</div>
<div id="counterContainer" style="padding-bottom: 10px;">
    <p class="counter-p"><%=ms_isonline %><span class="counter-detail"><%=application("activevisitors") %></span></p>
    <p class="counter-p"><%=ms_today %><span class="counter-detail"><%=readFile("/thongkehomnay.txt") %></span></p>
    <p class="counter-p"><%=ms_tongac %><span class="counter-detail"><%=readFile("/tongtruycap.txt") %></span></p>
</div>
<%end sub %>

<%sub video1() 
    sqlv = "Select * from V_news where categoryloai = 6"
    set rsv = server.CreateObject("ADODB.RECORDSET")
    rsv.Open sqlv,con,1

    if not rsv.EOF then  
%>
<div class="bot-20"></div>
<div class="other-header">

    <p style="">
        <%=getColVal("Newscategory","CategoryName","CategoryLoai = 6") %>
    </p>

</div>
<div class="list-pro-cate-container">
    <div class="list-pro-cate-video-slider">
        <%
                        do while not rsv.EOF

        %>
        <div class="hot-product-item list-pro-cate-item video">
            <div class="pro-rps" style="margin: auto;">
                <%
                   link =  rsv("description")
                   'Response.Write Left(link,4)
                    if Left(link,4) = "http" then
                %>
                <a href="<%=link %>">
                    <img style="width: 100%; height: 100%; max-width: 100%;" src="/images_upload/<%=getColVal("Picture","SmallPictureFileName","pictureId = '"&rsv("pictureId")&"'") %>">
                </a>

                <%
                    else
                    
                %>
                <iframe width="100%" height="100%" src="https://www.youtube.com/embed/<%=link %>" frameborder="0" allowfullscreen></iframe>
                <%end if%>
            </div>
        </div>
        <%
                       

                        rsv.MoveNext()
                        loop
                        set  rsv = nothing
                        
        %>
    </div>
    <div class="clearfix"></div>


    <div class="bot-20"></div>
</div>
<%
    
    end if
    end sub 
%>



<%sub search2(key) %>

<div class="container ">

    <div class="form-group" style="position: relative; height: 26px;">
        <div style="float: left; background: #75bb1f;" class="in-title">
            <h3 class="H-Title">Tìm kiếm</h3>
        </div>
    </div>
    <hr class="Hr-Title">

    <p class="clear search-txt "> Tìm kiếm với từ khóa :"<b style="color:#F00;"> <%=key%></b> "</p>

    <% 
    'response.write()
            call getSearch(key)
    %>
     <!-- <div class="clear" style="margin-bottom:100px;"></div> -->
    <br />

</div>
<%end sub %>




<%Function getColVal(tabName,colName,query) %> <!--getColVal("newscategory","categoryname","categoryid = '"&cid_&"'") Theo thu tu no se là bảng rồi đến tên mục, đến catergoryID -->
<% 
    if tabName = "" or colName = "" or query = "" then
        getColVal = ""
    else
        sqlCV = "Select top 1 "&colName&" from "&tabName&" where "&query 'chọn ra giá trị dòng đầu tiên cột categoryname; với điều kiện là có catergory ='807'chẳng hạn
        set rsCV = server.CreateObject("ADODB.RecordSet")
        rsCV.open sqlCV,con,1
            if not rsCV.EOF then
                getColVal = Trim(rsCV(colName))
            else
                getColVal = ""
            end if
        set rsCV = nothing
    end if
%>
<%End Function %>





<%Function Uni2NONE(sStr)
    Dim sTemp
    sTemp=Trim(sStr)
    
    'a
    sTemp=Replace(sTemp,"á","a")
    sTemp=Replace(sTemp,"à","a")
    sTemp=Replace(sTemp,"ả","a")
    sTemp=Replace(sTemp,"ã","a")
    sTemp=Replace(sTemp,"ạ","a")
    
    'ă
    sTemp=Replace(sTemp,"ă","a")
    sTemp=Replace(sTemp,"ắ","a")
    sTemp=Replace(sTemp,"ằ","a")
    sTemp=Replace(sTemp,"ẳ","a")
    sTemp=Replace(sTemp,"ẵ","a")
    sTemp=Replace(sTemp,"ặ","a")
    
    'â
    sTemp=Replace(sTemp,"â","a")
    sTemp=Replace(sTemp,"ấ","a")
    sTemp=Replace(sTemp,"ầ","a")
    sTemp=Replace(sTemp,"ẩ","a")
    sTemp=Replace(sTemp,"ẫ","a")
    sTemp=Replace(sTemp,"ậ","a")
    
    'đ
    sTemp=Replace(sTemp,"đ","d")

    
    'e
    sTemp=Replace(sTemp,"é","e")
    sTemp=Replace(sTemp,"è","e")
    sTemp=Replace(sTemp,"ẻ","e")
    sTemp=Replace(sTemp,"ẽ","e")
    sTemp=Replace(sTemp,"ẹ","e")
    
    'ê
    sTemp=Replace(sTemp,"ê","e")
    sTemp=Replace(sTemp,"ế","e")
    sTemp=Replace(sTemp,"ề","e")
    sTemp=Replace(sTemp,"ể","e")
    sTemp=Replace(sTemp,"ễ","e")
    sTemp=Replace(sTemp,"ệ","e")
    
    'i
    sTemp=Replace(sTemp,"í","i")
    sTemp=Replace(sTemp,"ì","i")
    sTemp=Replace(sTemp,"ỉ","i")
    sTemp=Replace(sTemp,"ĩ","i")
    sTemp=Replace(sTemp,"ị","i")
    
    'o
    sTemp=Replace(sTemp,"ó","o")
    sTemp=Replace(sTemp,"ò","o")
    sTemp=Replace(sTemp,"ỏ","o")
    sTemp=Replace(sTemp,"õ","o")
    sTemp=Replace(sTemp,"ọ","o")
    
    'ô
    sTemp=Replace(sTemp,"ô","o")
    sTemp=Replace(sTemp,"ố","o")
    sTemp=Replace(sTemp,"ồ","o")
    sTemp=Replace(sTemp,"ổ","o")
    sTemp=Replace(sTemp,"ỗ","o")
    sTemp=Replace(sTemp,"ộ","o")
    
    'ơ
    sTemp=Replace(sTemp,"ơ","o")
    sTemp=Replace(sTemp,"ớ","o")
    sTemp=Replace(sTemp,"ờ","o")
    sTemp=Replace(sTemp,"ở","o")
    sTemp=Replace(sTemp,"ỡ","o")
    sTemp=Replace(sTemp,"ợ","o")
    
    'u
    sTemp=Replace(sTemp,"ú","u")
    sTemp=Replace(sTemp,"ù","u")
    sTemp=Replace(sTemp,"ủ","u")
    sTemp=Replace(sTemp,"ũ","u")
    sTemp=Replace(sTemp,"ụ","u")
    
    'ư
    sTemp=Replace(sTemp,"ư","u")
    sTemp=Replace(sTemp,"ứ","u")
    sTemp=Replace(sTemp,"ừ","u")
    sTemp=Replace(sTemp,"ử","u")
    sTemp=Replace(sTemp,"ữ","u")
    sTemp=Replace(sTemp,"ự","u")
    
    'y
    sTemp=Replace(sTemp,"ý","y")
    sTemp=Replace(sTemp,"ỳ","y")
    sTemp=Replace(sTemp,"ỷ","y")
    sTemp=Replace(sTemp,"ỹ","y")
    sTemp=Replace(sTemp,"ỵ","y")
'---------------------------------Chữ hoa-------------------------------------------------
    'A
    sTemp=Replace(sTemp,"Á","A")
    sTemp=Replace(sTemp,"À","A")
    sTemp=Replace(sTemp,"Ả","A")
    sTemp=Replace(sTemp,"Ã","A")
    sTemp=Replace(sTemp,"Ạ","A")
    
    'Ă
    sTemp=Replace(sTemp,"Ă","A")
    sTemp=Replace(sTemp,"Ắ","A")
    sTemp=Replace(sTemp,"Ằ","A")
    sTemp=Replace(sTemp,"Ẳ","A")
    sTemp=Replace(sTemp,"Ẵ","A")
    sTemp=Replace(sTemp,"Ặ","A")
    
    'Â
    sTemp=Replace(sTemp,"Â","A")
    sTemp=Replace(sTemp,"Ấ","A")
    sTemp=Replace(sTemp,"Ầ","A")
    sTemp=Replace(sTemp,"Ẩ","A")
    sTemp=Replace(sTemp,"Ẫ","A")
    sTemp=Replace(sTemp,"Ậ","A")
    
    'Đ
    sTemp=Replace(sTemp,"Đ","D")
    
    'E
    sTemp=Replace(sTemp,"É","E")
    sTemp=Replace(sTemp,"È","E")
    sTemp=Replace(sTemp,"Ẻ","E")
    sTemp=Replace(sTemp,"Ẽ","E")
    sTemp=Replace(sTemp,"Ẹ","E")
    
    'Ê
    sTemp=Replace(sTemp,"Ê","E")
    sTemp=Replace(sTemp,"Ế","E")
    sTemp=Replace(sTemp,"Ề","E")
    sTemp=Replace(sTemp,"Ể","E")
    sTemp=Replace(sTemp,"Ễ","E")
    sTemp=Replace(sTemp,"Ệ","E")
    
    'I
    sTemp=Replace(sTemp,"Í","I")
    sTemp=Replace(sTemp,"Ì","I")
    sTemp=Replace(sTemp,"Ỉ","I")
    sTemp=Replace(sTemp,"Ĩ","I")
    sTemp=Replace(sTemp,"Ị","I")
    
    'O
    sTemp=Replace(sTemp,"Ó","O")
    sTemp=Replace(sTemp,"Ò","O")
    sTemp=Replace(sTemp,"Ỏ","O")
    sTemp=Replace(sTemp,"Õ","O")
    sTemp=Replace(sTemp,"Ọ","O")
    
    'Ô
    sTemp=Replace(sTemp,"Ô","O")
    sTemp=Replace(sTemp,"Ố","O")
    sTemp=Replace(sTemp,"Ồ","O")
    sTemp=Replace(sTemp,"Ổ","O")
    sTemp=Replace(sTemp,"Ỗ","O")
    sTemp=Replace(sTemp,"Ộ","O")
    
    'Ơ
    sTemp=Replace(sTemp,"Ơ","O")
    sTemp=Replace(sTemp,"Ớ","O")
    sTemp=Replace(sTemp,"Ờ","O")
    sTemp=Replace(sTemp,"Ở","O")
    sTemp=Replace(sTemp,"Ỡ","O")
    sTemp=Replace(sTemp,"Ợ","O")
    
    ''U
    sTemp=Replace(sTemp,"Ú","U")
    sTemp=Replace(sTemp,"Ù","U")
    sTemp=Replace(sTemp,"Ủ","U")
    sTemp=Replace(sTemp,"Ũ","U")
    sTemp=Replace(sTemp,"Ụ","U")
    
    ''Ư
    sTemp=Replace(sTemp,"Ư","U")
    sTemp=Replace(sTemp,"Ứ","U")
    sTemp=Replace(sTemp,"Ừ","U")
    sTemp=Replace(sTemp,"Ử","U")
    sTemp=Replace(sTemp,"Ữ","U")
    sTemp=Replace(sTemp,"Ự","U")
    
    ''Y
    sTemp=Replace(sTemp,"Ý","Y")
    sTemp=Replace(sTemp,"Ỳ","Y")
    sTemp=Replace(sTemp,"Ỷ","Y")
    sTemp=Replace(sTemp,"Ỹ","Y")
    sTemp=Replace(sTemp,"Ỵ","Y")

    'ký tự thừa
    sTemp=Replace(sTemp,"/","")
    sTemp=Replace(sTemp,"\","")
    sTemp=Replace(sTemp,",","")
    sTemp=Replace(sTemp,"&","")
    sTemp=Replace(sTemp,"$","")
    sTemp=Replace(sTemp,"~","")
    sTemp=Replace(sTemp,"*","")
    sTemp=Replace(sTemp,"#","")
    sTemp=Replace(sTemp,"(","")
    sTemp=Replace(sTemp,")","")
    sTemp=Replace(sTemp,"{","") 
    sTemp=Replace(sTemp,"}","")
    sTemp=Replace(sTemp,"|","")
    sTemp=Replace(sTemp,"'","''")
    sTemp=Replace(sTemp,"  ","")
    sTemp=Replace(sTemp," ","") 
        
    sTemp=replace(sTemp,"?","")
    sTemp=replace(sTemp,"%","phan-tram")
    
    Uni2NONE=sTemp
End Function
%>

<% function getLink(cateid,newsid,title) %>
<%
    link = ""
    if cateid <> "" then link = link & "/"&cateid
    if newsid <> "" then link = link & "/"&newsid
    if title <> "" then link = "/"& Replace(Uni2NONE(title)," ","-")&link&".html"
    if link <> "" then getLink = link else getLink = "#"
%>
<%end function %>

<%Sub LangIc() %>
<table class="table tb-ic" border="0">
    <tr>
        <td class="text-left mxh-ul-rs-none"><img src="/images/icon/hotline.png" class="img-responsive" style="margin-left: 0px; height: 35px;"></td>
        <td class="w-ic-lang"><img src="/images/icon/en.jpg" style="height: 25px; cursor: pointer; margin-top: 17px;" onclick="changeLang('0');" /></td>
        <td class="w-ic-lang"><img src="/images/icon/vn.jpg" style="height: 25px; cursor: pointer; margin-top: 17px;" onclick="changeLang('1');" /></td>
    </tr>
</table>
<%End Sub  %>

<%Sub image_home() %>
<%    
    sql1 = "SELECT TOP 6 * FROM NewsCategory WHERE  CategoryHome <> '0' And LanguageID='"&lang&"' ORDER BY CategoryHome "  'CateHome = -1 :đặc biệt

    set rs = Server.CreateObject("ADODB.Recordset")
    rs.open sql1,con,1
    item=1
    IF NOT rs.eof THEN 
    do while not rs.Eof
        CategoryImg =   Trim(rs("CategoryImg"))
        CategoryImg =   NewsImagePath&CategoryImg
        CategoryLink =  Trim(rs("CategoryLink"))
        if CategoryLink<>"" then
            linkuri=CategoryLink
        else
            linkuri = func_uri(rs("CategoryID"),"",rs("CategoryName"))
        end if
          
        if item=1 then
        Response.write "<div class='container' style =' margin-block: 2.5rem;'>"
            Response.write "<div class='row d-flex'>"
        end if        
             Response.write "<div class='col-md-2 col-sm-3 col-xs-3 text-center'>"&_
                "<a class='w3-text-black' href='"&linkuri&"'>"&_
                "<img style='max-width: 100%; width: 80px;' src='"&CategoryImg&"'>"&_
                "</a>"&_
                "<h3 class='image_home'><a class='w3-text-black' href='"&linkuri&"'>"&rs("CategoryName")&"</a></h3>"&_
                "</div>"
        if item=rs.recordcount then
            Response.write "</div><!---/.row--->"
            Response.write "</div><!---/.container--->"
        end if
    item=item+1
    rs.MoveNext
    Loop
    END IF
    rs.Close 
End Sub %>

<%Sub NewsLater() %>
<%
 sql = "SELECT TOP 4 * FROM V_News where LanguageID='"&lang&"' ORDER BY LastEditedDate DESC"
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1    
    if not rsi.eof then  
    item=1       
%>        
    <div class="news-later col-12 col-md-4 w3-padding-right">
        <div class="form-group css_relative">
        <div class="in-title">
            <h3 class="H-Title">BÀI VIẾT GẦN ĐÂY</h3>            
        </div>       
        </div>   
        <%
		link=""
        do while Not rsi.EOF
            Title       =   Trim(rsi("Title"))
            'Title       =   ucase(mid(Title,1,1))+lcase(mid(Title,2))
            desc_short  =   Trim(rsi("DecsBannerImage"))                           
            if  Len(desc_short)  > 130  then 
                desc_short = Left(desc_short,130)&"..."
            end if
            LastEditedDate  =   rsi("LastEditedDate")
            nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")             
            link = func_uri(rsi("CategoryID"),rsi("NewsID"),rsi("Title"))   
        %>     
        <div class="Item-news">
            <a href="<%=link %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=link %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <span class="text-justify"><%=desc_short %></span>
            </p>
            <%if item<rsi.RecordCount then  %>
            <hr />
            <%end if %>
        </div>
        <%
            item=item+1
            rsi.MoveNext
            Loop
        %>
        <%
            if IDC>0 then
                'Call Fs_regeditEmail() 
            end if
        %>
    </div>
<%
    set rsi = nothing
    end if 'end if 
End Sub
%>

<%Sub Fs_regeditEmail() %>      
<div class="form-group css_relative">
    <div class="in-title"><h3 class="H-Title">THEO DÕI BẢN TIN</h3></div>       
</div>   
<p>Theo dõi bản tin của chúng tôi để biết tin tức và bài viết mới nhất.Hãy cập nhật!</p>
<form name="Fremail" id="Fremail" method="post">
    <div class="form-group"> 
        <input class="form-control" id="inputName" name="inputName" placeholder="Họ và tên"/> 
    </div>
    <div class="form-group">  
        <input class="form-control" id="inputEmail" name="inputEmail" placeholder="Email"> 
    </div>
    <div class="form-group">  
        <input class="form-control btn btn-primary" id="btn_submit" name="btn_submit" type="button" value="Gửi"> 
    </div>
</form>
<script type="text/javascript">
    $("#btn_submit").click(function () {
        if ($('#inputName').val() == '') {
            $('#inputName').focus();
            swal("BQT", "Xin vui lòng nhập họ tên.");
        }
        else if ($('#inputEmail').val() == '') {
            $('#inputEmail').focus();
            swal("BQT", "Xin vui lòng nhập email.");
        }
        else if (!isEmail($('#inputEmail').val())) {
            $('#inputEmail').focus();
            swal("BQT", "Sai định dạng email.vd: abc@gmail.com");
        }
        else {
            Func_resEmail('send-email', '0');
        }
    });

    function isEmail(email) {
        var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;
        return regex.test(email);
    }
</script>
<% 
End Sub
%>

<%
Sub Fs_CateOther()
    sql = "SELECT TOP 3 * FROM NewsCategory WHERE  CategoryHome <> '0' And LanguageID='"&lang&"' ORDER BY CategoryHome "  'CateHome = -1 :đặc biệt
    'response.write sql
    set rs = Server.CreateObject("ADODB.Recordset")
    rs.open sql,con,1
    If not rs.eof Then  
    Response.write "<div class='container'>"    
    do while not rs.Eof            
        cid_        =   Trim(rs("CategoryID"))
        CateName    =   Trim(rs("CategoryName"))
        CateNote    =   Trim(rs("CategoryNote"))
        CateLoai    =   Trim(rs("CategoryLoai"))
        CategoryID  =   Trim(rs("CategoryID"))          
        linkcate    =   func_uri(cid_,"",rs("CategoryName"))   
    %>
    <div class="news-later col-md-4 col-sm-12 col-xs-12 w3-padding-right">
        <div class="form-group css_relative">
        <div class="in-title">
            <a href="<%=linkcate %>"> <h3 class="H-Title"><%=CateName %></h3></a>                       
        </div>       
        </div>   
        <%
        sql_item = "SELECT TOP 2 * FROM V_News where CategoryID ='"&cid_&"' and  LanguageID='"&lang&"' ORDER BY LastEditedDate DESC"
        set rsi = server.CreateObject("ADODB.RECORDSET")
        rsi.open sql_item,con,1    
        if not rsi.eof then
        item=1  
        do while Not rsi.EOF
            Title       =   Trim(rsi("Title"))
            Title       =   ucase(mid(Title,1,1))+lcase(mid(Title,2))
            desc_short  =   Trim(rsi("DecsBannerImage"))                           
            if  Len(desc_short)  > 130  then 
                desc_short = Left(desc_short,130)&"..."
            end if
            LastEditedDate  =   rsi("LastEditedDate")
            nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")             
            linkuri = func_uri(rsi("CategoryID"),rsi("NewsID"),rsi("Title"))                  
        %>     
        <div class="Item-news">
            <a href="<%=linkuri %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
            </p>
        </div>
        <%
            if item<rsi.RecordCount then    
                Response.write "<hr/>"
            end if
        item=item+1
        rsi.MoveNext
        Loop
        end if
        %>
    </div>      
    <%    
    rs.MoveNext
    Loop
    Response.write "</div><!---/.container--->"
    End If
    rs.Close
End Sub
%>

<%'---------------------------------------------------------------------------------------------------------------------------------------------------- %>
<%
Sub NewsHome()
    'Call image_home()
    sql = "SELECT TOP 10 * FROM NewsCategory WHERE  CategoryHome <> '0' And LanguageID='"&lang&"' ORDER BY CategoryHome "  'CateHome = -1 :đặc biệt
    set rsCollections = Server.CreateObject("ADODB.Recordset")
    rsCollections.open sql,con,1
    IF NOT rsCollections.eof THEN      
    ' Display :  interface 
    '1 : intro
    '2 : service
    '3 : product
    do while not rsCollections.Eof            
        cid_        =   Trim(rsCollections("CategoryID")) ' lấy ra id của cột
        CateName    =   Trim(rsCollections("CategoryName"))
        CateNote    =   Trim(rsCollections("CategoryNote"))
        CateLoai    =   Trim(rsCollections("CategoryLoai")) ' cai này để phân loại có 1 và 4
        CategoryID  =   Trim(rsCollections("CategoryID"))   ' lấy ra id của cột

    '===============================================
    SELECT CASE CateLoai 
        'CASE 4 ' Tin  tức
            'Call Ineter_F(CategoryID,4) 'Dùng cho xuống dưới phần body    
        CASE 2 ' San Pham
            Call Ineter_F(CategoryID,2)
            call Getonlybestsells(851,lang)
            response.Write("<div class='container sm-Banner-Ads'>")
            Call write_Ads2(IDCate,lang,7,0,0)
            response.Write("</div>")
        'CASE 3 ' Dich vu 2
            'Call Ineter_F(CategoryID,3)     
        CASE 1 ' Giới thiệu
            Call Ineter_F(CategoryID,1)  'Dùng cho lên đầu để giới thiệu  nhân vật các kiểu phần body 
        'Call Tindacbiet()  
        CASE 8 ' Tin  tức
            Call Ineter_F(CategoryID,8) ' cateId = catogoryLoai là nó tính theo loại 
        CASE 11 ' Tin  tức
                Call Ineter_F(CategoryID,11)
    END SELECT
    rsCollections.MoveNext
    Loop
    END IF
    rsCollections.Close
End Sub
%>
<%'---------------------------------------------------------------------------------------------------------------------------------------------------- %>

<%Function GetListParentCat(CatId)
    'Get Tree List CategoryId of Inpute Category.
    'Result is a string of CategoryId separated by spacebar, not include Input Category
    Dim i,ArrValue(100)
    i=0
    Dim rs1
    set rs1=Server.CreateObject("ADODB.Recordset")

      PCatId=CatId
      Do while PCatId<>0
        sql_GetListParentCat="select ParentCategoryId from NewsCategory where CategoryId=" & PCatId
        rs1.open sql_GetListParentCat,con,1
            PCatId=Cint(rs1("ParentCategoryId"))
            if PcatId<>0 then
                i=i+1
                ArrValue(i)=rs1("ParentCategoryId")
            end if
        rs1.close
      Loop
    GetListParentCat=Trim(Join(ArrValue))
End Function%>

<%'------------------------------------------------------------------------------------------------------------------------------------ %>
<%sub Fs_video(cid_,lgid_) 
    sql = " SELECT Title,Description,CategoryName,ParentCategoryID,NewsID,Author,PictureID,url_video FROM  V_News WHERE status = 4  And  url_video <> '' And   CategoryID = '"&cid_&"'  And LanguageID='"&lgid_&"' ORDER BY  CreationDate DESC "
    set rs = Server.CreateObject("ADODB.RECORDSET")
    rs.Open sql,con,1
    IF not rs.eof THEN
    'CName = Trim(rs("CategoryName"))
    'CName =  getColVal("newscategory","CategoryName"," CategoryID = '"&cid_&"' and LanguageID='"&lang&"'")
        'do while  not rs.EOF
    

%>

<div class="container">
    <h3 class="H-toppic">Video </h3>
    <hr class="Hr-toppic">
    <div class="test">
        <%
          Do while not rs.EOF


          v_title = Trim(rs("Title"))
          urlvideo = Trim(rs("url_video"))

          f_bd = "frameborder='0'  allowfullscreen "         
          str_ = InStr(1,urlvideo,"https://www.youtube.com/watch?v=") 'xác định link youtube. Nó start từ 1 và trả về số mà nó bắt đầu ví dụ như h ở vị trí 11
          IF urlvideo <> "" And   str_ > 0 THEN 
                idvd_ = Trim(Replace(urlvideo,"https://www.youtube.com/watch?v="," "))
          END IF
          
          w_ = "100%"
        %>
        <div class="P-5">

            <br />
            <iframe width="<%=w_ %>" height="" src="https://www.youtube.com/embed/<%=idvd_ %>" <%=f_bd %>></iframe>
            <span><%=v_title %></span>
        </div>

        <%
          rs.MoveNext
          Loop
        %>
    </div>
</div>
<%

    END IF
    rs.close
    end sub 
%>



<%
    sub Fs_Category(CateID_,LGID_)
    sql = "SELECT  * FROM  NewsCategory  WHERE LanguageID='"&LGID_&"' And ParentCategoryID = '"&CateID_&"'    And CategoryStatus  = '1'  or  CK ='1'   order by  CategoryOrder ASC  " 

    set rsCat = Server.CreateObject("ADODB.Recordset")
        rsCat.open sql,con,1
        IF NOT rsCat.eof THEN
%>
<%
             DO While NOT  rsCat.eof
                Cid_      = Trim(rsCat("CategoryID"))
                cateName    = Trim(rsCat("CategoryName"))
                CateLoai    = Trim(rsCat("CategoryLoai"))  
%>
<h3 class="text-link-cate"><%=cateName %> </h3>
<%

                 SELECT CASE CateLoai
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE "1" ' intro
                           Call F_intro(Cid_)
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE "2" ' news
                           Call F_news(Cid_)
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE "4" ' product
                           Call F_product(Cid_)
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE "8" ' services
                           Call F_product(Cid_)
                    CASE "10" ' services
                           Call support()
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE "11" ' product
                           Call F_video(Cid_)
'----------------------------------------------------------------------------------------------------------------------------------------------------
                    CASE ELSE ' intro
                           Call F_intro(Cid_)
'----------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------

                
                        
                    END SELECT   
             rsCat.MoveNext
             Loop        
%>
<%
    END IF
    rsCat.Close
    end sub 
%>


<%'---------------------------------------------------------------------------------------------------------------------------------------------------- %>


<%sub F_video(cid_) 
    sql = " SELECT Title,Description,CategoryName,ParentCategoryID,NewsID,Author,PictureID,url_video FROM  V_News WHERE status = 4  And  url_video <> '' And   CategoryID = '"&cid_&"' And AdsNews = '1' ORDER BY  CreationDate DESC "
    set rs = Server.CreateObject("ADODB.RECORDSET")
    rs.Open sql,con,1
    IF not rs.eof THEN
    'CName = Trim(rs("CategoryName"))
    'CName =  getColVal("newscategory","CategoryName"," CategoryID = '"&cid_&"' and LanguageID='"&lang&"'")
        'do while  not rs.EOF
    

%>
<%
          Do while not rs.EOF


          v_title = Trim(rs("Title"))
          urlvideo = Trim(rs("url_video"))

          f_bd = "frameborder='0'  allowfullscreen "         
          str_ = InStr(1,urlvideo,"https://www.youtube.com/watch?v=") ' xác định link youtube.
          IF urlvideo <> "" And   str_ > 0 THEN
                idvd_ = Trim(Replace(urlvideo,"https://www.youtube.com/watch?v="," "))
          END IF
          
          w_ = "100%"
%>
<div class="P-5">

    <br />
    <iframe width="<%=w_ %>" height="" src="https://www.youtube.com/embed/<%=idvd_ %>" <%=f_bd %>></iframe>
    <span><%=v_title %></span>
</div>

<%
          rs.MoveNext
          Loop
%>
<%

    END IF
    rs.close
    end sub 
%>


<%'---------------------------------------------------------------------------------------------------------------------------------------------------- %>

<%
    Function Is_Mobile()
        Set Regex = New RegExp
        With Regex
          .Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad)"
          .IgnoreCase = True
          .Global = True
        End With
        Match = Regex.test(Request.ServerVariables("HTTP_USER_AGENT"))
        If Match then
          Is_Mobile = True
        Else
          Is_Mobile = False
        End If
    End Function
%>


