<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
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
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
</head>
<body>
    <div class="container-fluid">
        <%
            Call header() 
        %>
    </div>
    <div class="container-fluid">
        <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <form method="post" name="fSearch" action="?">
                <table class="table  table-bordered">
                    <tr>
                        <th>TÌM KIẾM TIN & SẢN PHẨM </th>
                    </tr>
                    <tr>
                        <td>Tìm kiếm:&nbsp;			
                    <%
                        key_search=Trim(Request.Form("keyword"))
                        key_status=CInt(Request.Form("StatusId"))
                        CatID=0
	                    Ngay1=Day(now()-7)
	                    Thang1=Month(now()-7)
	                    Nam1=2004
	                    Ngay2=Day(now())
	                    Thang2=Month(now())
	                    Nam2=Year(now())

	                    if request.Form("action")="Search" then
	                    	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	                    	Thang1=GetNumeric(Request.form("Thang1"),0)
	                    	Nam1=GetNumeric(Request.form("Nam1"),0)
	                    	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	                    	Thang2=GetNumeric(Request.form("Thang2"),0)
	                    	Nam2=GetNumeric(Request.form("Nam2"),0)
                        else
                            Ngay1=Day(now()-7)
	                        Thang1=Month(now()-7)
	                        Nam1=2004
	                        Ngay2=Day(now())
	                        Thang2=Month(now())
	                        Nam2=Year(now())
                        end if
			            Call List_Date_WithName(Ngay1,"DD","Ngay1")
			            Call List_Month_WithName(Thang1,"MM","Thang1")
			            Call List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
                    %>
                            <img src="../images/right.jpg" width="9" height="9" />
                            <%
			            Call List_Date_WithName(Ngay2,"DD","Ngay2")
			            Call List_Month_WithName(Thang2,"MM","Thang2")
			            Call List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
                            %>    	</td>
                    </tr>
                    <tr>
                        <td class="form-horizontal">
                            <div class="form-group">
                                <div class="col-md-8">
                                    <input name="keyword" type="text" value="<%=Replace(key_search,"""","&quot;")%>" placeholder="Tiêu đề" class="form-control col-md-5">
                                </div>
                                <div class="col-md-2">
                                    <select name="StatusId" class="form-control">
                                        <option <% if key_status = 0  then Response.Write("Selected=""Selected""")%> value="0">Trạng thái</option>
                                         <option <% if key_status = 2  then Response.Write("Selected=""Selected""")%> value="2">Lưu</option>
                                        <option <% if key_status = 4  then Response.Write("Selected=""Selected""")%> value="4">Đăng</option>
                              
                                    </select>
                                </div>
                                <div class="col-md-2">
                                    <input type="submit" name="ButtonSearch" id="ButtonSearch" class="btn btn-primary" value="Tìm kiếm" />
                                </div>
                            </div>
                            <input type="hidden" name="action" value="Search">
                        </td>
                    </tr>
                </table>
            </form>

            <!--<%
                keyword=Trim(Request.Form("keyword"))
                ac_=Trim(Request.Form("action"))
                IF ac_ <> "" THEN
                
                IF IsEmpty(keyword) THEN
                    Top  = 20
                END IF
                lang = Session("Language")
                if lang = "" then lang = "VN"
	            	keywordold = keyword
	            	
	            	if keyword ="" then
	            		keyword	=	Request.QueryString("keyword")                    
	            	end if
	            	Keyword=Replace(keyword,"'","''")
	            	CategoryLoai=Request.Form("CategoryLoai")
	            	if CategoryLoai="" then
	            		CategoryLoai=Request.QueryString("CategoryLoai")
	            	end if         
	            	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
	            	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2	
                            
	            	StringSearch=UCase(Trim(Keyword))
	            	Set rs=server.CreateObject("ADODB.Recordset")
	            	sql1="SELECT "&Top&" * " &_
	            		"FROM	V_News where LanguageID = '"&lang&"' "
	            	sql1= sql1+ " and  (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) "
                    if StringSearch <> "" then
                        sql1= sql1+ "and  UPPER(Title) like N'%"&StringSearch&"%'"      '(UPPER(IDcode) = '"&StringSearch&"' or
                    end if
                    StatusId = Request.Form("StatusId")
                    IF StatusId <> "0"  THEN
                        sql1= sql1+ " AND status = '"&StatusId&"'"
                    END IF
                    sql1= sql1 + " ORDER BY CreationDate DESC"
	                set rs=server.CreateObject("ADODB.Recordset")
	                rs.open sql1,con,1  
                    'response.write sql1
	                if not rs.eof then
            %>

            <table class="table table-bordered w3-hoverable">
                <tr class="w3-blue">
                    <td>TT</td>
                    <th>Tiêu đề tin</th>
                    <td>Chuyên</td>
                    <td>Lượt xem</td>
                    <th>Xửlý</th>
                </tr>
                <%
                    Do while not rs.eof 
						link_url=func_uri(rs("Categoryid"),rs("newsid"),rs("title"))  
                %>
                <tr>
                    <td><%=stt%>.</td>
                    <td><a target="_blank" href="<%=link_url%>"><%=rs("Title")%></a></td>
                    <td><%=GetListParentCatNameOfCatId(rs("CategoryId"))%></td>
                    <td><%=rs("NewsCount")%></td>
                    <td>
                        <a href="news_addedit.asp?iStatus=edit&newsid=<%=rs("NewsId")%>&catid=<%=rs("Categoryid")%>">
                            <img src="/administrator/images/icon_edit_topic.gif" width="15" height="15" border="0" title="Sửa"></a>
                        <a href="javascript: winpopup('/administrator/news/news_delete.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',400,220);">
                            <img src="/administrator/images/icon_closed_topic.gif" width="15" height="15" border="0" title="Xóa"></a>
                    </td>
                </tr>
                <%
	                stt=stt+1
                    rs.movenext
                    Loop
                %>
            </table>
            <%	
	            'end if 'if not rs.eof then
	            rs.close
	            set rs=nothing
                else            
                end if              
                END IF             
            %>-->
<!--!-- viet lai ma khong can an tim kiem -->
    <%
        lang = Session("Language")
                if lang = "" then lang = "VN"
	            	keywordold = keyword
	            	
	            	if keyword ="" then
	            		keyword	=	Request.QueryString("keyword")                    
	            	end if
	            	Keyword=Replace(keyword,"'","''")
	            	CategoryLoai=Request.Form("CategoryLoai")
	            	if CategoryLoai="" then
	            		CategoryLoai=Request.QueryString("CategoryLoai")
	            	end if         
	            	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
	            	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2	
                            
	            	StringSearch=UCase(Trim(Keyword))
	            	Set rs=server.CreateObject("ADODB.Recordset")
	            	sql1="SELECT "&Top&" * " &_
	            		"FROM	V_News where LanguageID = '"&lang&"' "
	            	sql1= sql1+ " and  (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) "
                    if StringSearch <> "" then
                        sql1= sql1+ "and  UPPER(Title) like N'%"&StringSearch&"%'"      '(UPPER(IDcode) = '"&StringSearch&"' or
         
                    end if
                    'StatusId = Request.Form("StatusId")
                   'IF StatusId <> "0"  THEN
                       'sql1= sql1+ " AND status = '"&StatusId&"'"
                   'END IF
                    sql1= sql1 + " ORDER BY CreationDate DESC"
                    
	                set rs=server.CreateObject("ADODB.Recordset")
	                rs.open sql1,con,1  
                    'response.write sql1
	                if not rs.eof then
        %>
     <table class="table table-bordered w3-hoverable">
                <tr class="w3-blue">
                    <td>TT</td>
                    <th>Tiêu đề tin</th>
                    <td>Chuyên</td>
                    <td>Lượt xem</td>
                    <th>Xửlý</th>
                </tr>
                <%
                    Do while not rs.eof 
						link_url=func_uri(rs("Categoryid"),rs("newsid"),rs("title"))  
                %>
                <tr>
                    <td><%=stt%>.</td>
                    <td><a target="_blank" href="<%=link_url%>"><%=rs("Title")%></a></td>
                    <td><%=GetListParentCatNameOfCatId(rs("CategoryId"))%></td>
                    <td><%=rs("NewsCount")%></td>
                    <td>
                        <a href="news_addedit.asp?iStatus=edit&newsid=<%=rs("NewsId")%>&catid=<%=rs("Categoryid")%>">
                            <img src="/administrator/images/icon_edit_topic.gif" width="15" height="15" border="0" title="Sửa"></a>
                        <a href="javascript: winpopup('/administrator/news/news_delete.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',400,220);">
                            <img src="/administrator/images/icon_closed_topic.gif" width="15" height="15" border="0" title="Xóa"></a>
                    </td>
                </tr>
                <%
	                stt=stt+1
                    rs.movenext
                    Loop
                %>
            </table>
            <%	
	            'end if 'if not rs.eof then
	            rs.close
	            set rs=nothing
                else            
                end if              
                          
            %>
        </div>
    </div>
   
    <%Call Footer()%>
</body>
</html>
