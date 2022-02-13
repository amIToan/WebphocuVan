<%@  language="VBSCRIPT" codepage="65001" %>
<%Call PhanQuyen("QLyHeThong")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
    lang = Session("Language")
    if lang = "" then lang = "VN"
CategoryLoai	=	request.QueryString("CategoryLoai")
if isnumeric(CategoryLoai) = false or CategoryLoai = "" then
	CategoryLoai	=	 -1
end if

if Trim(request.Form("action"))<>"" and isnumeric(request.Form("CatId")) then
	action=Trim(replace(request.Form("action"),"'","''"))
	catid=Cint(request.Form("CatId"))
	catlevel=Cint(Request.Form("Catlevel"))
	
	Call MoveCat(Action,CatId,CatLevel,lang)
	Call Update_PrentCategoryId(lang)
	Call Update_YoungestChildren(lang)
	
End if%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script language="JavaScript1.2" src="/administrator/inc/common.js"></script>
    <script language="JavaScript">
        function fCatSubmit(actionvalue, catid, catlevel) {
            document.fCat.action.value = actionvalue;
            document.fCat.CatId.value = catid;
            document.fCat.CatLevel.value = catlevel;
            document.fCat.submit();
        }
    </script>
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>
<body>
    <div class="container-fluid">
        <%Call header()%>
    </div>
    <div class="container-fluid">
        <div class="col-md-2" style="background: #001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <table class="table tab-content">
                <tr>
                    <th>DANH SÁCH CHUYỂN MỤC</th>
                </tr>
                <tr>
                    <td>                    
                        <a style="font-weight:bold;color:blue;" href="javascript: winpopup('cat_addnew.asp','<%=lang%>&CategoryLoai=<%=CategoryLoai%>',500,500);">TẠO CHUYÊN MỤC MỚI</a>
                    </td>
                </tr>
            </table>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM NewsCategory where languageId='" & lang & "' "
	sql	=	sql + " ORDER BY CategoryOrder"
	rs.open sql,con,1
            %>
            <table class="table table-bordered w3-hoverable">
                <tr align="center" bgcolor="FFFFFF">
                    <td><strong></strong><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
      chuyên mục</font></strong></td>
                    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Loại 
      </font></strong></td>
                    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Thuộc</font></strong></td>
                    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vị trí</font></strong></td>
                    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Di 
      chuyển</font></strong></td>
                    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
                </tr>
                <form action="<%=Request.ServerVariables("Script_Name")%>?lang=<%=lang%>&CategoryLoai=<%=CategoryLoai%>" method="post" name="fCat">
                    <input type="hidden" name="action" value="">
                    <input type="hidden" name="CatId" value="">
                    <input type="hidden" name="CatLevel" value="">
                </form>
                <%boo=false
  Do while not rs.eof
  	boo=not boo
  	Response.Write "<tr "
	if boo then
    	response.Write "bgcolor=""#E6E8E9"">" & vbNewline
	else
		response.Write "bgcolor=""#FFFFFF"">" & vbNewline
	end if

	
   	Response.Write "	<td>"
	Response.Write "	<font size=""2"" face=""Arial, Helvetica, sans-serif"">"
	for i=2 to Cint(rs("CategoryLevel"))
		response.Write("&nbsp;&nbsp;")
	Next 
	if Cint(rs("YoungestChildren"))=0 and Cint(rs("CategoryLevel"))<>1 then
		response.write "&nbsp;-&nbsp;" & rs("CategoryName") & "<br>"
	else
		if Cint(rs("CategoryLevel"))=1 then
			response.write "&nbsp;&#8226;&nbsp;<b>" & rs("CategoryName") & "</b><br>"
		else
			response.write "&nbsp;&#8226;&nbsp;" & rs("CategoryName") & "<br>"
		end if
	end if
    
	Response.Write	"</font></td>" & vbNewline
    Response.Write	"<td align=""center""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & GetNameOfCategoryStatus(Cint(rs("CategoryStatus"))) & "</font></td>" & vbNewline
    Response.Write "<td align=""center""><font size=""2"" face=""Arial, Helvetica, sans-serif"">&nbsp;" & GetNameOfCategoryLoai(rs("CategoryLoai")) & "</font></td>" & vbNewline &_
    				"<td align=""center"">" & vbNewline &_					
						"<a href=""javascript: fCatSubmit('left'," & rs("CategoryId") & "," & rs("CategoryLevel") & ");"" style=""text-decoration: none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif"" color=""#000000"">Tr&#225;i</font></a>" & vbNewline &_
						"<a href=""javascript: fCatSubmit('up'," & rs("CategoryId") & "," & rs("CategoryOrder") & ");"" style=""text-decoration: none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif"" color=""#000000"">L&#234;n</font></a>" & vbNewline &_
						"<a href=""javascript: fCatSubmit('down'," & rs("CategoryId") & "," & rs("CategoryOrder") & ");"" style=""text-decoration: none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif"" color=""#000000"">Xu&#7889;ng</font></a>" & vbNewline &_
						"<a href=""javascript: fCatSubmit('right'," & rs("CategoryId") & "," & rs("CategoryLevel") & ");"" style=""text-decoration: none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif"" color=""#000000"">Ph&#7843;i</font></a>" & vbNewline &_
					"</td>" & vbNewline
      CategoryHome = ""

      select case Cint(rs("CategoryHome"))
           case "0"
                CategoryHome = "Không đặt"
           case "-1"
                CategoryHome = "Đặc biệt"
           case else
                CategoryHome = rs("CategoryHome")
      end select
     
     
    Response.Write	"<td align=""center"">"&_
                        "<font size=""2"" face=""Arial, Helvetica, sans-serif"">" & CategoryHome & "</font>"&_
                    "</td>" & vbNewline
	Response.Write	"<td align=""center"">" & vbNewline &_					
						"<a class='w3-btn w3-blue w3-round' href=""javascript: winpopup('cat_edit.asp','" & lang & "&Catid=" & rs("CategoryID") & "',600,500);"" style=""text-decoration:none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif""><i class='fa fa-pencil-square-o' aria-hidden='true'></i> S&#7917;a</font></a>" & vbNewline &_
						"<a class='w3-btn w3-red w3-round' href=""javascript: winpopup('cat_delete.asp','" & lang & "&Catid=" & rs("CategoryID") & "',300,220);"" style=""text-decoration:none"">" & vbNewline &_
						"<font size=""1"" face=""Arial, Helvetica, sans-serif""><i class='fa fa-trash-o fa-lg' aria-hidden='true'></i> X&#243;a</font></a>" & vbNewline &_
					"</td>" & vbNewline &_
     		"</tr>" & vbNewline
  rs.movenext
  Loop
  rs.close
  set rs=nothing
                %>
            </table>
        </div>
    </div>
    <%Call Footer()%>
</body>
</html>
