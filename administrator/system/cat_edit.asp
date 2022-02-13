<%@  language="VBSCRIPT" codepage="65001" %>
<%Call PhanQuyen("QLyHeThong")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	 lang = Session("Language")
    if lang = "" then lang = "VN"
	
	CatId=Request.QueryString("Catid")
	if not isNumeric(catid) then
		response.End()
	end if
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
if Trim(Request.Querystring("action"))="Update" then			
		sError=False
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SetMaxSize 10000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		set uploadCategoryImg = Upload.Files("CategoryImg")

		If uploadCategoryImg Is Nothing Then
			CategoryImgPath=""
		else
		    CategoryImgPaths = ""
            dem = 0
            For Each File in Upload.Files
            Filetype = Right(File.Filename,len(File.Filename)-Instr(File.Filename,"."))
		    if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg"  and Lcase(Filetype)<>"png" then
				sError=True
				sCategoryImgPath="Không là file ảnh"
		    else
                dem = dem + 1
		   		CategoryImgPath="IMGCategory_Path_"&dem&"_"& CatId & "." & Filetype
                Response.Write("CategoryImgPath"&CategoryImgPath)
		        if CategoryImgPath <>"" then 
			    File.SaveAs Path & "\" & CategoryImgPath
		        end if
                CategoryImgPaths =  CategoryImgPath&CategoryImgPaths
		    end if
            Next
		End If
           
		CatName=Trim(replace(Upload.form("CategoryName"),"'","''"))
		ParentCatId=Cint(Upload.form("CategoryId"))
		CatLink=Trim(replace(Upload.form("CategoryLink"),"'","''"))
		'Lang
		if Upload.form("CategoryStatus")<>"" and isnumeric(Upload.form("CategoryStatus")) then
			CategoryStatus=Cint(Upload.form("CategoryStatus"))
		else
			CategoryStatus=0
		end if
		if Upload.form("CategoryLoai")<>"" and isnumeric(Upload.form("CategoryLoai")) then
			CategoryLoai=Cint(Upload.form("CategoryLoai"))
		else
			CategoryLoai=0
		end if
		CatNote=Trim(replace(Upload.form("CategoryNote"),"'","''"))
		iChietKhau = Trim(replace(Upload.form("txtCK"),"'","''"))
		if isnumeric(iChietKhau) = false then
			iChietKhau = 0
		end if
		if Upload.form("cbChietKhau")<>"" and isnumeric(Upload.form("cbChietKhau")) then
			iEditCK = CLng(Upload.form("cbChietKhau"))
		else
			iEditCK = 0
		end if

		if Upload.form("cbMaLoai")<>"" and isnumeric(Upload.form("cbMaLoai")) then
			iMaLoai = CLng(Upload.form("cbMaLoai"))
		else
			iMaLoai = 0
		end if

        meta_keyword=Trim(replace(Upload.form("meta_keyword"),"'","''"))
        meta_desc=Trim(replace(Upload.form("meta_desc"),"'","''"))
   		
		Ngayphat=Trim(replace(Upload.form("Category_Ngayphat"),"'","''"))		
		Giophat=Trim(replace(Upload.form("Category_Giophat"),"'","''"))				
		Pre_ParentCatId=Cint(Upload.form("Pre_ParentCatId"))
		CategoryHome=Cint(Upload.form("CategoryHome"))

		if ParentCatId<>Pre_ParentCatId then
			sql="select CategoryOrder,CategoryLevel from NewsCategory where CategoryId=" & ParentCatId
			rs.open sql,con,1
			if not rs.eof then
			CatOrder=Cint(rs("CategoryOrder"))+1
			CatLevel=Cint(rs("CategoryLevel"))+1
			end if
			rs.close
	
			sql="Update NewsCategory set CategoryOrder = CategoryOrder + 1 Where languageid='" & lang & "' and CategoryOrder>='" & CatOrder&"'"

			rs.open sql,con,1
		end if
		
		YoungestChildren=0
		
		sql="Update NewsCategory set "
		sql=sql & "CategoryName=N'" & CatName & "'"
		if CategoryImgPath<>"" then	
			sql=sql & ",CategoryImg='" & CategoryImgPaths & "'"
		end if
		sql=sql & ",ParentCategoryId=" & ParentCatId
		if ParentCatId<>Pre_ParentCatId then
			sql=sql & ",CategoryLevel=" & CatLevel
			sql=sql & ",CategoryOrder=" & CatOrder
		end if
		sql=sql & ",YoungestChildren=" & YoungestChildren
		sql=sql & ",CategoryLink='" & CatLink & "'"
		sql=sql & ",CategoryNote=N'" & CatNote & "'"
		sql=sql & ",Category_Ngayphat=N'" & Ngayphat & "'"		
		sql=sql & ",Category_Giophat =N'" & Giophat & "'"	
		sql=sql & ",CategoryHome=" & CategoryHome 		
		sql=sql & ",CategoryStatus=" & CategoryStatus 
		sql=sql & ",CategoryLoai='" & CategoryLoai& "'"	
        sql=sql & ",meta_keyword=N'" & meta_keyword& "'"
		sql=sql & ",meta_desc=N'" & meta_desc& "'"	
		sql=sql & " where CategoryId=" & CatId
    Response.Write sql

		rs.open sql,con,1
		set rs=nothing

		
		Call Update_PrentCategoryId(lang)
		Call Update_YoungestChildren(lang)
		
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location=""cat_list.asp?lang=" & lang & "&CategoryLoai="&CategoryLoai&""";" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
		response.End()
	
end if	
	sql="select * from NewsCategory where CategoryId=" & CatId
	rs.open sql,con,1
		CatName=Trim(rs("CategoryName"))
		CategoryImg=Trim(rs("CategoryImg"))
		ParentCatId=Cint(rs("ParentCategoryId"))
		YoungestChildren =Cint(rs("YoungestChildren"))
		CatLink=Trim(rs("CategoryLink"))
		CatNote=Trim(rs("CategoryNote"))				
		CategoryStatus=Cint(rs("CategoryStatus"))
		CategoryHome=rs("CategoryHome")
		CategoryLoai=Cint(rs("CategoryLoai"))
		CategoryNote=rs("CategoryNote")
		Ck=rs("CK")
        meta_keyword=rs("meta_keyword")
        meta_desc=rs("meta_desc")
		CategoryLevel=Cint(rs("CategoryLevel"))
	rs.close
	set rs=nothing
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
</head>
<body>

    <div class="col-md-8 col-md-offset-1">
        <form name="fNew" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=Update&<%=Request.ServerVariables("QUERY_STRING")%>" enctype="multipart/form-data">
            <table border="0" class="table table-condensed">
                <tr>
                    <th colspan="2">CHỈNH SỬA CHUYÊN MỤC
                    </th>
                </tr>
                <tr>
                    <td style="width: 30%;">Chuyên mục:</td>
                    <td>
                        <input name="CategoryName" type="text" id="CategoryName" class="form-control" maxlength="100" value="<%=CatName%>"></td>
                </tr>
                <tr>
                    <td>Vị trí:</td>
                    <td>
                        <%Call List_Category_Name(ParentCatId,"CategoryID","Kh&#244;ng chuy&#234;n m&#7909;c n&#224;o",lang,-1)%>	  </td>
                </tr>
                <tr>
                    <td>Liên kết:</td>
                    <td>
                        <input name="CategoryLink" type="text" id="CategoryLink" class="form-control" maxlength="100" value="<%=CatLink%>"></td>
                </tr>
                <tr>
                    <td></td>
                    <td><%if CategoryImg<>"" then %>
                        <%
                           CategoryImgArr = Split(CategoryImg,"|")
                            if UBound(CategoryImgArr) >0 then
                                For i=0 to UBound(CategoryImgArr) -1
                        %>
                                <img style="max-width: 100%; height: 50px;" src="<%=NewsImagePath&CategoryImgArr(i)%>">
                        <%
                                Next
                            else
                        %>
                                <img style="max-width: 100%; height: 50px;" src="<%=NewsImagePath&CategoryImg%>">

                        <%
                             end if
                        %>


                        <%End if%></td>
                </tr>
                <tr>
                    <td>Img:</td>
                    <td>
                        <input name="CategoryImg" type="file" id="CategoryImg" multiple="multiple"></td>
                </tr>

                <tr>
                    <td>keywords:</td>
                    <td>
                        <input name="meta_keyword" type="text" id="meta_keyword" class="form-control" value="<%=meta_keyword %>"/></td>
                </tr>
                <tr>
                    <td>description:</td>
                    <td>
                        <input name="meta_desc" type="text" id="meta_desc" class="form-control" value="<%=meta_desc %>"></td>
                </tr>
                <tr>
                    <td>Ngôn ngữ: </td>
                    <td><b><%=GetNameOfLanguage(lang)%></b></td>
                </tr>
                <tr>
                    <td>Trạng thái: </td>
                    <td><%Call ListStatusOfCategory(CategoryStatus)%></td>
                </tr>
                <tr style="display: none;">
                    <td>Chuyên mục:</td>
                    <td>
                        <input name="cbo_viewcat" type="checkbox" value="1" <%=sck %> />
                        Hiển thị</td>
                </tr>
                <%'if CategoryLevel = 1 then%>
                <tr>
                    <td>Danh mục trang chủ </td>
                    <td>
                        <%
              sqltemp = "select count(CategoryID) as icount from NewsCategory where  LanguageID = '"& lang&"'"
              set rsTemp = Server.CreateObject("ADODB.recordset")
	          rsTemp.open sqltemp,con,1
              icount = rsTemp("icount") + 1
              set rstemp=nothing
                        %>
                        <select id="CategoryHome" name="CategoryHome" class="form-control" style="min-width: 200px;max-width: 200px;">
                            <option value="0" <%if CategoryHome=0 then%> selected <%end if %>>Không đặt</option>
                            <option value="-1" <%if CategoryHome=-1 then%> selected <%end if %>>Đặc biệt</option>
                            <%for i = 1 to icount %>
                            <option value="<%=i%>" <%if i = CategoryHome then%> selected <%end if %>><%=i %></option>
                            <%next %>
                        </select>

                    </td>
                </tr>
                <%'end if%>
                <tr>
                    <td>Giao diện:</td>
                    <td><%Call ListLoaiOfCategory(CategoryLoai)%></td>
                </tr>              
                <tr>
                    <td>Mô Tả:</td>
                    <td>
                        <textarea name="CategoryNote" class="form-control" cols="50" rows="5" id="CategoryNote"><%=CategoryNote%></textarea></td>
                </tr>

                <tr>
                    <td colspan="2" class="text-right">
                        <input type="submit" name="Submit" value=" Sửa " class="btn btn-primary">
                        <input type="button" name="Submit2" value="Đóng cửa sổ" onclick="javascript: window.close();" class="btn btn-primary" style="margin-left: 5px;">&nbsp;                   
                        <input type="hidden" name="Pre_ParentCatId" value="<%=ParentCatId%>">
                    </td>
                </tr>
            </table>
        </form>

    </div>
</body>
</html>
<script src='../inc/news.js'></script>
<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.forms[0].CategoryNote)</script>
<script language="javascript">
    function testCK() {
        if (document.fNew.CategoryLoai.value == 3) {
            document.fNew.txtCK.readonly = true;
        }
        else {
            document.fNew.txtCK.readonly = false;
        }
    }
</script>

