<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<!--#include virtual="/include/Fs_cotruct.asp"-->
  <% sub smallBanner()%>
    <div class="container py-3">
        <div class="sm-Banner">
        </div>
    </div>
<% End Sub%>
<% Sub Getonlybestsells(cid_,lang) 
    sql = "SELECT  * FROM V_News WHERE   status = '4' And  (CategoryID = '"&cid_&"')  And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'cid id ở đây là catergoryID
    set rsSell = server.CreateObject("ADODB.RECORDSET")
    Response.Write "<br />"
    rsSell.open sql,con,1
    if not rsSell.eof then   
    CName =  getColVal("newscategory","categoryname","categoryid = '"&cid_&"'") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
    link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html" 
  %>
<div class="container" style="margin-bottom:20px;"> 
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title  col-12 col-md-6 ">
            <a href="<%=link_cate %>"><h3 class="H-Title"><%=CName %></h3></a>        
        </div>       
    </div>
 <%
            item=1
            Do while Not rsSell.EOF
            tem_ = 0
            Title   = Trim(rsSell("Title"))
            desc    = Trim(rsSell("Author"))
            desc_short    = Trim(rsSell("DecsBannerImage"))
            DiscountedPrice = Trim(rsSell("PriceNet"))
            Price = Dis_str_money(Trim(rsSell("Price")))
            Unit = Trim(rsSell("Unit"))  
            if  Len(desc_short)  > 180  then 
                desc_short = Left(desc_short,180)&"..."
            end if

            cateID  = Trim(rsSell("CategoryID"))
            NewsID  = Trim(rsSell("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsSell("PictureId")&"'")   
            url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)
            
            if item=1 then
                Response.write "<div class='index-group owl-carousel owl-theme '>"
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
            if item = rsSell.recordcount then 
                response.Write("</div>") '<!---/.index-group--->
            end if
            item=item+1
            rsSell.MoveNext
            Loop
        %>
 <script type="text/javascript">
            var owl = $(".index-group");
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
    <% End If %>
<% End Sub%>
<% 
    Sub NewsandLibary()
%>
<div class="container">
    <div class="row">
        <% call News ()%>
        <% Call imageLiberary() %>
        <% Call NewsvideoLiberary() %>
    </div>
</div>
<% End Sub %>

<% sub News() 
    sql =" select * from V_News where  status = '4' And  (CategoryLoai = 4)  And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "
    set rsNews = server.CreateObject("ADODB.RECORDSET")
    response.Write("<br>")
    rsNews.open sql,con,1
    if not rsNews.eof then   
    CName =  getColVal("newscategory","categoryname","categoryLoai = 4") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
    cid_ =  getColVal("newscategory","categoryID","categoryLoai = 4")
    link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html"
    rsNews.PageSize = 8
    tem_ = 1
 %>
    <div class=" col-12 col-md-4" style="margin-bottom:20px;"> 
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title  ">
            <a href="<%=link_cate %>"><h3 class="H-Title"><%=CName %></h3></a>        
        </div>       
    </div>
       <% do while Not rsNews.EOF and tem_ < rsNews.PageSize
            Title   = Trim(rsNews("Title"))
            desc    = Trim(rsNews("Author"))
            desc_short    = Trim(rsNews("DecsBannerImage"))
            if  Len(desc_short)  > 30  then 
                desc_ = desc
            else
                desc_ = Left(desc_short,30)&"..."
            end if                
            cateID  = Trim(rsNews("CategoryID"))
            NewsID  = Trim(rsNews("NewsID"))
            nimg    =  getColVal("Picture","SmallPictureFileName","PictureId = '"&rsNews("PictureId")&"'")   
            url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title) 'để bay vào trang chi tiết 
%>
        <div class="Item-news "> <!--col-md-6 col-xs-12-->
            <a href="<%=linkuri %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <span class="text-justify add-text"><%=desc_short %></span>
            </p>
        </div>
        <hr>
        <%
            rsNews.MoveNext
            tem_ = tem_+1
            Loop
        %>
<%
    set rsNews = nothing
    end if 'end if dich vu 
%>
  </div>
<% End Sub %>
<% sub imageLiberary()
    sql = "SELECT * FROM V_News WHERE   status = '4' and (CategoryLoai = 3) order by LastEditedDate DESC " 
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1   
    if not rsi.eof then
        count = rsi.RecordCount 
        CName =  getColVal("newscategory","categoryname","categoryLoai = 3") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
        cid_ =  getColVal("newscategory","categoryID","categoryLoai = 3")
        link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html"
        'linkuri = func_uri(cid_,"",CName)
%>  
<div class=" col-12 col-md-4" style="margin-bottom:20px;">
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title  ">
            <a href="<%=link_cate %>"><h3 class="H-Title"><%=CName %></h3></a>        
        </div>       
    </div>
        <%
        i=0
        tem_ = 1
        do while Not rsi.EOF and  tem_ =< 6        
        Title       =   Trim(rsi("Title"))
        desc_short  =   Trim(rsi("DecsBannerImage"))                           
        if  Len(desc_short)  > 50 then 
            desc_short = mid(desc_short,1,50)&"..."
        end if
        LastEditedDate  =   rsi("LastEditedDate")
        nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
        linkuri = func_uri(rsi("CategoryID"),rsi("NewsID"),rsi("Title"))
        %>    
        <%if tem_=1 then %>
           <div class="">
             <a href="<%=linkuri %>">
                <img class="img-responsive-self" src="/images_upload/<%=nimg %>" alt=<%=Title%> />
            </a>
            <p>
                <a href="<%=linkuri %>" class=" w3-margin-left sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <!--<span class="w3-margin-left text-justify add-text"><%=desc_short %></span>-->
            </p>
        </div>
        <hr>
        <%else %>
        <div class="Item-news "> <!--col-md-6 col-xs-12-->
            <a href="<%=linkuri %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <!-- <span class="text-justify add-text"><%=desc_short %></span> -->
            </p>
        </div>
        <hr>
       <% end if %> 
   <%   
        tem_=tem_+1
        i = i + 1
        rsi.MoveNext
        Loop
      end if
      set rsi = nothing %>
     </div><!---/.col-md-4---> 
 <%End Sub %>
<% sub NewsvideoLiberary()
    sql = "SELECT * FROM V_News WHERE   status = '4' and (CategoryLoai = 7) order by LastEditedDate DESC " 
    set rsi = server.CreateObject("ADODB.RECORDSET")
    rsi.open sql,con,1   
    if not rsi.eof then
        count = rsi.RecordCount 
        CName =  getColVal("newscategory","categoryname","categoryLoai = 7") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
        cid_ =  getColVal("newscategory","categoryID","categoryLoai = 7")
        link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html"
%>  
<div class=" col-12 col-md-4" style="margin-bottom:20px;">
    <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title  ">
            <a href="<%=link_cate %>"><h3 class="H-Title"><%=CName %></h3></a>        
        </div>       
    </div>
        <%
        i=0
        tem_ = 1
        do while Not rsi.EOF and  tem_ =< 6        
        Title       =   Trim(rsi("Title"))
        desc_short  =   Trim(rsi("DecsBannerImage"))
        urlvideo = Trim(rsi("url_video"))
         f_bd = "frameborder='0'  allowfullscreen "
        str_ = InStr(1,urlvideo,"https://www.youtube.com/watch?v=") 'xác định link youtube. Nó start từ 1 và nó tìm xem chuỗi đó có nó chứa hay không và vị trí bắt đầu chuỗi str2
            IF urlvideo <> "" And   str_ > 0 THEN 
                idvd_ = Trim(Replace(urlvideo,"https://www.youtube.com/watch?v="," "))
          END IF
          w_ = "100%"
        if  Len(desc_short)  > 50 then 
            desc_short = mid(desc_short,1,50)&"..."
        end if
        LastEditedDate  =   rsi("LastEditedDate")
        nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsi("PictureId")&"'")   
        linkuri = func_uri(rsi("CategoryID"),rsi("NewsID"),rsi("Title"))
        %>    
        <%if tem_=1 then %>
           <div class="">
             <a href="<%=linkuri %>">
                <iframe width="<%=w_ %>" style="aspect-ratio:4/3" height="" src="https://www.youtube.com/embed/<%=idvd_ %>" <%=f_bd %>></iframe>
            </a>
            <p>
                <a href="<%=linkuri %>" class=" w3-margin-left sl-news-Tl add-text"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <!--<span class="w3-margin-left text-justify add-text"><%=desc_short %></span>-->
            </p>
        </div>
        <hr>
        <%else %>
        <div class="Item-news "> <!--col-md-6 col-xs-12-->
            <a href="<%=linkuri %>"><img class="img-responsive Av-news2" src="/images_upload/<%=nimg %>" alt=<%=Title%> /></a>
            <p>
                <a href="<%=linkuri %>" class="sl-news-Tl"><%=Title %></a>
                <span class="news-date"><%=GetFullDateVn(LastEditedDate) %></span>
                <span class="text-justify add-text"><%=desc_short %></span>
            </p>
        </div>
        <hr>
       <% end if %> 
   <%   
        tem_=tem_+1
        i = i + 1
        rsi.MoveNext
        Loop
      end if
      set rsi = nothing %>
     </div><!---/.col-md-4--->
<% End Sub %>
<% Sub Getfeedback() 
      sql = "SELECT  * FROM V_News WHERE   status = '4' And  (CategoryLoai = '5')  And LanguageID='"&lang&"'  Order By   LastEditedDate DESC "   'cid id ở đây là catergoryID
    set rsSell = server.CreateObject("ADODB.RECORDSET")
    Response.Write "<br />"
    rsSell.open sql,con,1
    if not rsSell.eof then   
    CName =  getColVal("newscategory","categoryname","categoryLoai = 5") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
    cid_ =  getColVal("newscategory","categoryID","categoryLoai = 5")
    link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&cid_&".html" 
%>
  <div class="container" style="margin-bottom:20px;">
     <div class="form-group" style="position: relative;height: 36px;">
        <div class="in-title col-lg-12 col-md-12 col-sm-6 col-xs-12">
            <a href="<%=link_cate %>" class="d-flex"><h3 class="H-Title" style="margin:auto"><%=CName %></h3></a>        
        </div>       
    </div>
    <div class="text-center px-3 pb-2"><i>Phản hồi của những khách hàng đã và đang sử dụng sản phẩm trong suốt những năm qua</i></div>
<%
            item=1
            Do while Not rsSell.EOF
            tem_ = 0
            Title   = Trim(rsSell("Title"))
            author    = Trim(rsSell("Author"))
            desc_short    = Trim(rsSell("DecsBannerImage"))
            Price = Trim(rsSell("Price"))
            Unit = Trim(rsSell("Unit"))  
            if  Len(desc_short)  > 180  then 
                desc_short = Left(desc_short,180)&"..."
            end if

            cateID  = Trim(rsSell("CategoryID"))
            NewsID  = Trim(rsSell("NewsID"))
            nimg    =       getColVal("Picture","SmallPictureFileName","PictureId = '"&rsSell("PictureId")&"'")   
            url_img = "style= 'background:url(images_upload/"&nimg&")"&";'" 
            linkuri = func_uri(cateID,NewsID,Title)
            
            if item=1 then
                Response.write "<div class='feedback owl-carousel owl-theme '>"
            end if                    
       %>
    <div class="px-3 py-3">
        <div class="border border-rounded border-success pt-4">
                <a href="<%=linkuri %>"><img src="/images_upload/<%=nimg %>" class="d-block mx-auto image-avatar"></a>
                <div class="news-content w3-padding">
                    <p class="news-title">
                    <b class="w3-text-black"><%=Title %> </b>
                    </p>
                    <div class="news-desc "><%=desc_short %></div>
                    <div class="news-desc "><%=author %></div>
                    <div class="news-desc w3-text-green"><% if Price <> "" and Price <> 0 then response.Write(Price&"<span class='News-unit'>"&Unit&"</span>")%></div>
                    <button type="button" class="d-block btn btn-success mx-auto my-3"><a href="<%=linkuri %>" style="color:white">Đọc thêm </a></button>
                 </div>
           </div>
     </div>
<%if item = rsSell.recordcount then 
                response.Write("</div>") '<!---/.index-group--->
            end if


            item=item+1
            rsSell.MoveNext
            Loop
        %>
 <script type="text/javascript">
     var owl = $(".feedback ");
     //console.log(owl);
     owl.owlCarousel({
         items: 3,
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
                 items: 3
             }
         }
     });
 </script>
 </div>
    <% End If %>
<% End Sub%>
<!-- đây là hàm test cho trang sản phẩm chưa đưa vào -->
<%Sub commonProducts(Cate_,cLoai_,Lgid_)
	    sqlGetproducts = "SELECT * FROM V_News WHERE   status = '4' And  CategoryID = '"&Cate_&"' or ParentCategoryID ="&Cate_&"   And (url_video IS NULL or url_video ='')   And LanguageID='"&Lgid_&"' Order by LastEditedDate DESC "    
		set rsProducts = server.CreateObject("ADODB.RECORDSET")
		rsProducts.open sqlGetproducts,con,1
		rsProducts.PageSize = 10
		pagecount=rsProducts.pagecount
        pageperbook=10
		IF Not rsProducts.EOF   THEN 'IF 1 
        CName  =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
%>
    <div class="container pdb-40">
        <div class="form-group">
            <div class="in-title css_relative">
                <h3 class="H-Title"><%=CName %></h3>
            </div>
            <div class="site-map"><i>Trang chủ &nbsp; <i class="fa fa-angle-double-right" aria-hidden="true"></i>&nbsp; <%=CName %> </i></div>
     </div>
	 <div class="row">
		<div class="col-12 col-md-3 pt-5 w3-hide-small">
            <% call leftMenuShowcat(Cate_,cLoai_,Lgid_) %>
		</div>     
		<div class="col-12 col-md-9"> 
            <div class="row">
            <div class="col-12 d-flex pb-3">
                <div class="col-6 ">...............</div>
                <div class="col-6 d-flex align-items-center justify-content-end"> 
                    <label for="" class="col-3 col-form-label">Sắp xếp :</label>
                    <select class="col-7 form-control">
                        <option>Theo mặc định</option>
                        <option>Theo giá</option>
                        <option>Mua nhiều</option>
                        <option>Theo chữ cái</option>
                    </select>
                </div>
            </div>
		<%
		if request.Querystring("page")<>"" then
			    page=Clng(request.Querystring("page"))
		    else
			    page=1
		    end if

		      rsProducts.AbsolutePage = CLng(page)

         j=0
         i = 1
		do while not rsProducts.eof and j<rsProducts.pagesize 
        Title   = Trim(rsProducts("Title"))
        desc_short    = Trim(rsProducts("DecsBannerImage"))
        Ncontent    = Trim(rsProducts("Body")) 
        cateID  = Trim(rsProducts("CategoryID"))
        NewsID  = Trim(rsProducts("NewsID"))
        nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rsProducts("PictureId")&"'")    
        linkuri = func_uri(cateID,NewsID,Title)              
        CName =  getColVal("newscategory","categoryname","categoryid = '"&Cate_&"'")
		Price = Dis_str_money(Trim(rsProducts("Price")))
        Unit = Trim(rsProducts("Unit"))  
        if  Len(desc_short)  > 180  then 
             desc_short = Left(desc_short,180)&"..."
        End if
		linkuri = func_uri(cateID,NewsID,Title)
%>
		<div class="col-12 col-md-4 px-3">
			<div class="item w3-border w3-round-xxlarge">
                <figure class="product-image index-group"> 
                <a href="<%=linkuri %>"><img src="/images_upload/<%=nimg %>" class="img-responsive mx-auto d-block"></a>
                    <div class="product-over">
                                        <button
                                            class="btn btn-small addToCart"
                                            data-product-id="2"
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
                    <a href="<%=linkuri %>" class="w3-text-black"><%=Title %> </a>
                </p>
                    <span class="stars"></span>
                    <div class="news-desc "><%=desc_short %></div>
                    <div class="news-desc w3-text-green w3-large"><% if Price <> "" and Price <> 0 then response.Write(Price&"<span class='News-unit'>"&Unit&"</span>")%></div>
                 </div>
            </div>
		</div>
	<%
		i=i+1
        j=j+1
		rsProducts.MoveNext
		Loop
	%>
            <div class="col-12 d-flex pb-3">
                <%Call phantrang(page,pagecount,pageperbook) %>
            </div>
			</div>
		</div>
	</div>
	<% End IF
		set rsProducts = nothing
	%>
   </div>     
<%ENd Sub%>

<!-- Ham left cho showcart -->
<% sub leftMenuShowcat(catergoryId,cLoai,Lang)  %>
            <%
                sqlMenu="SELECT    * " &_
                             "FROM      NewsCategory " &_
                             "WHERE   LanguageId = '"&lang&"' and  (Categoryloai ="&cLoai&") and CategoryLevel = 1 ORDER BY CategoryOrder"
                set rsMenu = server.CreateObject("ADODB.RECORDSET")
                rsMenu.open sqlMenu,con,1
                if not rsMenu.eof then   
                CName =  getColVal("newscategory","categoryname","categoryid = '"&catergoryId&"'") 'sẽ lấy đc ra catergoryName mà có id= catergoryID
                link_cate= "/"&Replace(Uni2NONE(CName)," ","")&"/"&catergoryId&".html" 
                %>
            <nav class="nav-product-menu ">
                <h4 class="left-header">
                    <% =CName %>
                </h4>
            <ul class="left-product-menu">
                <% 
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
					Call childMenuShowcat(YoungestChildren,CategoryID,varLink,CategoryName,active)     
				rsMenu.MoveNext()
				loop
                    %>
            </ul>
        <script type="text/javascript" langguage="javascript">
            // click to show under the chevron1
            const navProducts = document.querySelector(".left-product-menu");
            navProducts.removeChild(navProducts.lastElementChild)
            const h4Title = document.querySelector(".left-header");
            h4Title.style.width = navProducts.clientWidth + "px";
            navProducts.addEventListener("click", (e) => {
                e.preventDefault;
                const currentTarget = e.target.closest(".fa-chevron-down");
                const showUlMoblie = currentTarget.nextElementSibling;
                const childrenFlop = Array.from(showUlMoblie.children);
                const arrow = currentTarget.firstElementChild;
                if (!currentTarget) return;
                childrenFlop.forEach((element, index) => {
                    element.style.transition = `all 0.7s ease ${index / 30}s`;
                    element.classList.toggle("transitionForLi");
                });
                showUlMoblie.classList.toggle("height-side-menu");
                arrow.classList.toggle("bx-chevron-down-reverse")

            })
        </script>
        </nav>
<% 
    End if
    End Sub%>

<!-- sub-menu-showcart -->
      <% sub childMenuShowcat(YoungestChildren,CategoryID,varLink,CategoryName,active)

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
                   "WHERE     (ParentCategoryID = '"&CategoryID&"') and (CategoryStatus = 2 or CategoryStatus = 4 )  ORDER BY CategoryOrder"
          
            set rsSM = Server.CreateObject("ADODB.RECORDSET")
            rsSM.Open sqlSM,con,1
%><li class="nav__item">
   
    <%
            IF  not rsSM.eof THEN
    %>
	<a href="<%=varLink %>" style="cursor:pointer;" class ="nav__link dropdown__link <%=active %>"><%=CategoryName %></a>
     <i class="fas fa-chevron-down dropdown__icon"></i>
    <ul class="dropdown__menu " role="menu">
        <%
            do while not rsSM.eof

   
            call childMenuShowcat(rsSM("YoungestChildren"),rsSM("CategoryID"),getLink(rsSM("CategoryID"),"",rsSM("CategoryName")),rsSM("CategoryName"),"")
         
            rsSM.MoveNext()
            loop
        %>
    </ul>
    <% else%>
	 <a href="<%=varLink %>"><%=CategoryName %></a>
     <%END IF %>
</li>
<hr class="left-product-menu-hr">

<%
    else
%>
<li class="nav__item <%=active %>"><a href="<%=varLink %>" ><%=CategoryName %></a></li>
<hr class="left-product-menu-hr">
<%
    end if
    
%>

<% end sub %>

<!-- trang bán chi tiết -->
<%Sub sliderImageL(title,PictureId) %>
<div class="d-block">
<% 
redim PictureFile(16)
redim ContentPicture(16)
sql="select * From Picture Where PictureId=" & PictureId
Set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,con,1				
    if not rs.EOF then
		SmallPictureFileName    =   rs("SmallPictureFileName")
		LargePictureFileName    =   rs("LargePictureFileName")
		For n=1 to 16
			PictureFile(n)      =	rs("PictureFile"&n)
			ContentPicture(n)   =   rs("ContentPicture"&n)
		Next
    end if
rs.close 
Set rs=nothing       
%>
<div class="d-block">
    <a href="<%=NewsImagePath&SmallPictureFileName %>" data-fancybox="images">
        <img src="<%=NewsImagePath&SmallPictureFileName %>" class="lazy" alt="<%=title %>">
    </a>
</div>
<%For i =1 to UBound(PictureFile)  %>
    <%if PictureFile(i)<>"" then %>
    <div class="item item-media">
    <a href="<%=NewsImagePath&PictureFile(i) %>" data-fancybox="images">
        <img src="<%=NewsImagePath&PictureFile(i) %>" class="lazy" alt="<%=ContentPicture(i) %>">
    </a>
    </div>
    <%end if %>
<%Next %>
</div>
<%End Sub %>

<%Sub Fs_NewsDetail_Product(NewsID,cLoai) 
    sqlN = "SELECT * FROM V_News WHERE   status = '4' And  NewsID = '"&NewsID&"'"
    set rsN = Server.CreateObject("ADODB.Recordset")
    rsN.open sqlN,con,1
    If not rsN.eof then
        Title       = Trim(rsN("Title"))
        desc        = Trim(rsN("Author"))
        FTitle      = Trim(rsN("Description")) 
        Ncontent    = Trim(rsN("Body")) 
        cateID      = Trim(rsN("CategoryID"))
        NewsID      = Trim(rsN("NewsID"))
        PictureID   = Trim(rsN("PictureID"))
        if rsN("PriceNet") <> ""  then
                originPrice = Clng(Trim(rsN("Price")))
                discountpercent = Clng(trim(rsN("PriceNet"))) 
                Price = Clng(rsN("Price")) - ((discountpercent * Clng(rsN("Price"))) / 100 )
                Price = Dis_str_money(Price)
        else
                Price = Dis_str_money(Trim(rsN("Price")))
        End IF 
        Price       = Dis_str_money(Trim(rsN("Price")))
        desc_short  = Trim(rsN("DecsBannerImage"))
        Unit = Trim(rsN("Unit"))
        CName       =  getColVal("newscategory","categoryname","categoryid = '"&cateID&"'")
        urlvideo = Trim(rsN("url_video"))
         f_bd = "frameborder='0'  allowfullscreen "
        str_ = InStr(1,urlvideo,"https://www.youtube.com/watch?v=") 'xác định link youtube. Nó start từ 1 và nó tìm xem chuỗi đó có nó chứa hay không và vị trí bắt đầu chuỗi str2
            IF urlvideo <> "" And   str_ > 0 THEN 
                idvd_ = Trim(Replace(urlvideo,"https://www.youtube.com/watch?v="," "))
        END IF
        w_ = "100%"
    redim PictureFile(16)
    redim ContentPicture(16)
    sqlPicture="select * From Picture Where PictureId="&PictureID
    Set rs=server.CreateObject("ADODB.Recordset")
    rs.open sqlPicture,con,1				
        if not rs.EOF then
		    SmallPictureFileName    =   rs("SmallPictureFileName")
		    LargePictureFileName    =   rs("LargePictureFileName")
		    For n=1 to 8
			PictureFile(n)      =	rs("PictureFile"&n)
			ContentPicture(n)   =   rs("ContentPicture"&n)
		Next
    end if
    rs.close 
    Set rs=nothing       
%>
<link href="/flickity/flickity.min.css" rel="stylesheet" />
<link href="/flickity/fullscreen.css" rel="stylesheet" />
<link href="/product-details/style.css" rel="stylesheet" />
    <div class="container">
        <%=GetListParentCatNameOfCatId2(cateID,NewsID)%>
    </div>
<div class="container pb-5">
    <div class="row">
        <div class="col-12 col-md-9">
             <div class = "card-wrapper">
      <div class = "card product-under">
        <!-- card left -->
        <div class = "  product-imgs">
          <div class="carousel carousel-main " data-flickity='{"prevNextButtons": false,"fullscreen": true, "lazyLoad": 1 }'>
            <%For i =1 to UBound(PictureFile)  %>
                    <%if PictureFile(i)<>"" then %>
                         <div class="carousel-cell">
                            <img data-flickity-lazyload="<%=NewsImagePath&PictureFile(i) %>" class="lazy carousel-cell-image" alt="<%=ContentPicture(i) %>">\
                    </div>
                    <%end if %>
                <%Next %>
          </div>
          <div class="carousel carousel-nav"
                    data-flickity='{ "asNavFor": ".carousel-main", "contain": true, "pageDots": false, "prevNextButtons": false}'>
                    <%For i =1 to UBound(PictureFile)  %>
                    <%if PictureFile(i)<>"" then %>
                       <div class="carousel-cell">
                        <img src="<%=NewsImagePath&PictureFile(i) %>" class="d-block lazy" alt="<%=ContentPicture(i)%>" >
                       </div>
                    <%end if %>
                    <%Next %>
                  </div>
        </div>
        <!-- card right -->
        <div class = "product-content">
          <h2 class = "product-title productName"><% =Title %></h2>
          <a href = "/" class = "product-link">Tại cửa hàng Tôilàthảomộc</a>
          <div class = "product-rating">
            <i class = "fas fa-star"></i>
            <i class = "fas fa-star"></i>
            <i class = "fas fa-star"></i>
            <i class = "fas fa-star"></i>
            <i class = "fas fa-star-half-alt"></i>
            <span>4.7(21)</span>
          </div>

          <div class = "product-price">
            <p class = "last-price">Giá cũ: <span><% if originPrice <> "" and originPrice <> 0 then response.Write(originPrice&"<span class='News-unit'>"&Unit&"</span>")%></span></p>
            <p class = "new-price" > Giá đã giảm:
             <% if Price <> "" and Price <> 0 then response.Write("<span class='priceValue'>"&Price&"</span>"&"<span class='News-unit '>"&Unit&"</span>")%>
            </p>
            
          </div>

          <div class = "product-detail">
            <h3>About this item: </h3>
            <p> <% =desc_short %></p>
            <ul>
              <li>Color: <span>----</span></li>
              <li>Tình trạng: <span>Còn hàng</span></li>
              <li>Loại: <span>----</span></li>
              <li>Giao hàng: <span>Toàn quốc</span></li>
              <li>Phí ship (tham khảo): <span>Tùy địa điểm</span></li>
            </ul>
          </div>

          <div class = "purchase-info">
            <input type = "number" min = "1" value = "1">
            <button type = "button" class = "btn addToCart" data-product-id="<%=NewsID%>">
              Add to Cart <i class = "fas fa-shopping-cart"></i>
            </button>
            <button type = "button" class = "btn">Compare</button>
          </div>

          <div class = "detail-social-links">
            <div>Share At: </div>
            <a href = "https://www.facebook.com/" target="_blank">
              <i class = "fab fa-facebook-f"></i>
            </a>
            <a href = "https://twitter.com/?lang=vi">
              <i class = "fab fa-twitter"></i>
            </a>
            <a href = "https://www.instagram.com/">
              <i class = "fab fa-instagram"></i>
            </a>
            <a href = "https://web.whatsapp.com/">
              <i class = "fab fa-whatsapp"></i>
            </a>
            <a href = "https://www.pinterest.com/">
              <i class = "fab fa-pinterest"></i>
            </a>
          </div>
        </div>
      </div>
    </div>
            <button class="btn btn-success my-2 btn-click"> Mô tả</button>
            <div class="border p-3 rounded toggle-content">
                <div class="intro-body" ><%= %></div>
                <% if idvd_ <> "" then%>
                    <div class="d-flex justify-content-center pb-3">
                        <iframe width="<%=w_ %>" style="aspect-ratio:4/3" height="" src="https://www.youtube.com/embed/<%=idvd_ %>" <%=f_bd %>></iframe>
                    </div>
                <%End if%>
            </div>
        </div>
        <div class="col-12 col-md-3">
            <% call siderbar(15,cateID) %>
        </div>
    </div>
			<div class="fb-comments" data-href="https://toilathaomoc.com/" data-width="" data-numposts="10"></div>
</div>
    
 
<script type="text/javascript">
    $(document).ready(function () {
    $(".btn-click").click(function(){
       $(".toggle-content").toggle(1000);
  });
});
</script>       
<script src="/flickity/flickity.pkgd.min.js"></script>
<script src="/flickity/fullscreen.js"></script>
<%
    END IF 'not eof
    rsN.Close
%>
<%End Sub %>
<!--------------------buying function-------------------------------->
<% Sub Orderdetails() %>
    <link href="/interfaces/css/getandbuy.css" rel="stylesheet" />
    <div class="producstOnCart hide">
                <div class="overlay"></div>
                <div class="top">
                    <button id="closeButton">
                        <i class="fas fa-times-circle"></i>
                    </button>
                    <h3>Giỏ hàng</h3>
                </div>
                <ul id="buyItems"> 
				</ul>
                <button class="btn checkout" onclick="window.location.href='/cartlists'">
                <!-- <a href="/cartlists" style="color: white"> -->
                Check out
                </a>
                </button>
    </div>
<% End Sub %>

<%Sub SearchByTitle(Keyword)
    If Request.QueryString("cid") = "timkiem.html" and Keyword <> "" Then
        Set rs = Server.CreateObject("ADODB.RecordSet")
        SQL = "select * from V_News where Title like N'%"&Keyword&"%' order by LastEditedDate desc"
        rs.Open SQL, con ,1,3
        Session("Search_Text") = Keyword
    ElseIf Request.QueryString("cid") <> "timkiem.html" and Session("Search_Text") <> "" Then
        Set rs = Server.CreateObject("ADODB.RecordSet")
        SQL = "select * from V_News where Title like N'%"&Session("Search_Text")&"%' order by LastEditedDate desc"
        rs.Open SQL, con ,1,3
    ElseIf Keyword = "" and Request.QueryString("cid") = "timkiem.html" Then
        Session("Search_Text") = ""
        Set rs = Server.CreateObject("ADODB.RecordSet")
        SQL = "select * from V_News  order by LastEditedDate desc"
        rs.Open SQL, con,1,3
    End if

    If not rs.eof Then
        Dim PageLen,PageNo,TotalRecord,TotalPage,No,intID
        PageLen = 12
		TotalRecord = rs.RecordCount
		rs.PageSize = PageLen
		TotalPage = rs.PageCount
        pagination = 2
        if request.Querystring("page")<>"" then
			    page=Clng(request.Querystring("page"))
		    else
			    page=1
		    end if
		rs.AbsolutePage = CLng(page)

%>
    <div class="container">
            <div class="row mb-3">
                <h4 class="text-start">
                    <%
                    If Session("Search_Text") = "" Then
                        Response.Write ("Tất cả các kết quả :")
                    Else
                        Response.Write ("Tìm Kiếm Với Từ Khóa : "&Session("Search_Text"))
                    End if
                    %>
                </h4>
            </div>
            <div class="row">
                <%
                No = 1
                Do While not rs.eof and No <= PageLen
                NewsID  = Trim(rs("NewsID"))
                author  = Trim(rs("Author"))
                desc_short    = Trim(rs("DecsBannerImage"))
                if  Len(desc_short)  > 180  then 
                desc_short = Left(desc_short,180)&"..."
                end if
                nimg    = getColVal("Picture","SmallPictureFileName","PictureId = '"&rs("PictureId")&"'")
                LinkDetail= "/"&Replace(Uni2NONE(rs("Title"))," ","-")&"/"&rs("CategoryID")&"/"&rs("NewsID")
                %>
                    <div class="col-12 col-md-3 mb-3">
                        <div class="border border-rounded border-success pt-4">
                            <div class="px-3 py-3">
                            <a href="<%=LinkDetail %>"><img src="/images_upload/<%=nimg %>" class="d-block mx-auto image-avatar"></a>
                            <div class="news-content w3-padding">
                    <p class="news-title">
                    <b class="w3-text-black"><%=Title %> </b>
                    </p>
                    <div class="news-desc "><%=desc_short %></div>
                    <div class="news-desc "><%=author %></div>
                    <button type="button" class="d-block btn btn-success mx-auto my-3"><a href="<%=LinkDetail %>" style="color:white">Đọc thêm </a></button>
                 </div>
                        </div>
                        </div>
                    </div>
                <%
                No = No + 1
                rs.movenext
                Loop%>
            </div>
           <% call phantrang( page,TotalPage,pagination) %> 
        </div>
<%  Else 
    response.write("<h4 class='container text-start'> Không có kết quả phù hợp với tìm kiếm của bạn</h4>")
    End if
    set rs = nothing
End Sub%>

<%Sub Item_support_Toan(Lang)%>
<!-- support online -->
    <link href="/interfaces/css/getandbuy.css" rel="stylesheet" />
    <div id="style-switcher-toggle" class="off"></div>
    <div class="supportContainer hide">
        <div class="top">
                    <button id="closeButton2">
                        <i class="fas fa-times-circle"></i>
                    </button>
                    <h3>Hỗ trợ phần mềm</h3>
        </div>
<%
     Set rsSup=server.CreateObject("ADODB.Recordset")
     sql_support="SELECT * FROM SupportYahoo "
     rsSup.open sql_support,con,1
     If not rsSup.eof Then
	    Do while not rsSup.Eof 
            sidZalo=   rsSup("idzalo")   
            sHoTen  =   rsSup("HoTen") 
            sGhiChu =   rsSup("GhiChu")  
            sMobile =   rsSup("Mobile")  
            sEmail  =   rsSup("Email")
            sPicture=   rsSup("Picturepath")
    %> 
    <div class="w3-panel sup-item w3-border-bottom">
    <div class="w3-left sup-info w3-left-align " style="width: 60%">
        <h4><% =sHoTen %></h4>
        <p>Hỗ trợ phần mềm</p>
        <p><a rel="noopener" href="Tel: <% =sMobile%>"><i class="fa fa-volume-control-phone text_bluec"></i><% =sMobile%></a></p>
        <p><a rel="noopener" href="mailto:<% =sEmail%>"><i class="fa fa-envelope-o text_bluec" ></i><% =sEmail%></a></p>  
        <p><a rel="noopener" href="https://zalo.me/<% =sidZalo%>"><img class="w3-circle w3-image" src="/images/icons/zalo.png" / alt="<%=sidZalo%>" width="30" > Zalo</a></p>  
    </div>
    <div class="w3-right sup-img" style="width: 40%">
        <br />
      <img src="<%=NewsImagePath&sPicture%>" alt="" class="w3-circle w3-image" width="120"/>     
    </div>
    </div> 
    <%
	rsSup.movenext
	Loop
    End If 'end if not rsSup.EOF
    rsSup.close
	Set rsSup = nothing
    %>
</div>
<% End sub %>








