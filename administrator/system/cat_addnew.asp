<%@  language="VBSCRIPT" codepage="65001" %>
<%Call PhanQuyen("QLyHeThong")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	lang=Request.QueryString("param")
	lang=replace(lang,"'","''")
	CategoryLoai	=	GetNumeric(Request.QueryString("CategoryLoai"),-1)
if Request.Querystring("action")="Insert" then
	sError=False
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		CategoryID=GetMaxId("NewsCategory", "CategoryID", "")
		
		set uploadCategoryImg = Upload.Files("CategoryImg")
		If uploadCategoryImg Is Nothing Then
			CategoryImgPath=""
		else
		   Filetype = Right(uploadCategoryImg.Filename,len(uploadCategoryImg.Filename)-Instr(uploadCategoryImg.Filename,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" then
				sError=True
				sCategoryImgPath="Không là file ảnh"
		   else
		   		CategoryImgPath="IMGCategory_Path_" & CategoryID & "." & Filetype
		   end if
		End If
		if CategoryImgPath <>"" then uploadCategoryImg.SaveAs Path & "\" & CategoryImgPath

	if Trim(Upload.form("CategoryName"))<>"" and isnumeric(Upload.form("CategoryId")) then
		Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		CatName=Trim(replace(Upload.form("CategoryName"),"'","''"))

		cbo_viewcat=Trim(replace(Upload.form("cbo_viewcat"),"'","''"))
        IF cbo_viewcat <> "" THEN 
            viewcat =  1
        ELSE 
            viewcat =  0
        END IF
    
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
		if isCategoryNote(CatNote) = true and CatNote <>"" then	
%>
<script language="javascript">
    history.back();
    alert('Ma loai da ton tai xin moi ban nhap lai!')
</script>
<%
			Response.End()
		end if
		if ParentCatId=0 then
			sql="select max(CategoryOrder) as maxOrder from NewsCategory"
			rs.open sql,con,1
				if IsNull(rs("maxOrder")) then
					CatOrder=1
				else
					CatOrder=Cint(rs("maxOrder"))+1
				end if
			CatLevel=1
			rs.close
		else
			sql="select CategoryOrder,CategoryLevel from NewsCategory where CategoryId=" & ParentCatId
			rs.open sql,con,1
			CatOrder=Cint(rs("CategoryOrder"))+1
			CatLevel=Cint(rs("CategoryLevel"))+1
			rs.close
			
			sql="Update NewsCategory"
			sql=sql & " SET CategoryOrder=CategoryOrder+1"
			sql=sql & " Where LanguageId='" & lang & "' and CategoryOrder>=" & CatOrder

			rs.open sql,con,1
			
		end if
		YoungestChildren=0
		
		
		sql="insert into NewsCategory (CategoryName,CategoryImg,ParentCategoryId,CategoryLevel,CategoryOrder,"
		sql=sql & "YoungestChildren,CategoryLink,LanguageId,CategoryNote,CategoryStatus,CK,CategoryHome,CategoryLoai) values "
		sql=sql & "(N'" & CatName & "'"
		sql=sql & ",'" & CategoryImgPath& "'"
		sql=sql & "," & ParentCatId
		sql=sql & "," & CatLevel
		sql=sql & "," & CatOrder
		sql=sql & "," & YoungestChildren
		sql=sql & ",'" & CatLink & "'"
		sql=sql & ",'" & lang & "'"
		sql=sql & ",N'" & CatNote & "'"
		sql=sql & "," & CategoryStatus
        sql=sql & "," & viewcat
		sql=sql & ",'0'"
		sql=sql & "," & CategoryLoai & ")"	
		rs.open sql,con,1
		set rs=nothing
		Call Update_PrentCategoryId(lang)
		Call Update_YoungestChildren(lang)
		
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location=""cat_list.asp?lang=" & lang & "&CategoryLoai="& CategoryLoai &""";" & vbNewline &_
			"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	end if
End if	
%>
<html>

<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<script src="/ckeditor/ckeditor.js" type="text/javascript"></script>
<script src="/ckfinder/ckfinder.js" type="text/javascript"></script>


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form name="fNew" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>&action=Insert" enctype="multipart/form-data">
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
            <tr align="center" valign="middle">
                <td height="40" colspan="2" valign="middle"><strong>Chuyên 
        mục mới</strong></td>
            </tr>
            <tr>
                <td>Chuyên  mục:</td>
                <td>
                    <input name="CategoryName" type="text" id="CategoryName" size="30" maxlength="100">
                    <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong></font>)</font></td>
            </tr>
            <tr>
                <td>Vị trí sau:</td>
                <td>
                    <%
		Call List_Category(0, "Cu&#7889;i danh s&#225;ch",lang,-1)
                    %>	  </td>
            </tr>
            <tr>
                <td>Liên kết:</td>
                <td>
                    <input name="CategoryLink" type="text" id="CategoryLink" size="30" maxlength="100"></td>
            </tr>
            <tr>
            <tr>
                <td>Img:</td>
                <td>
                    <input name="CategoryImg" type="file" id="CategoryImg" size="30" maxlength="100"></td>
            </tr>
            <tr>
                <td>Ngôn ngữ:</td>
                <td><b><%=GetNameOfLanguage(lang)%></b></td>
            </tr>
            <tr>
                <td>Trạng thái:</td>
                <td><%Call ListStatusOfCategory(0)%></td>
            </tr>
            <tr>
                <td>Chuyên mục:</td>
                <td>
                    <input name="cbo_viewcat" type="checkbox" value="1" />
                    Hiển thị</td>
            </tr>
            <tr>
                <td>Loại:</td>
                <td><%Call ListLoaiOfCategory(0)%></td>
            </tr>

            <tr>
                <td align="center" colspan="2" height="35" valign="bottom">
                    <textarea name="CategoryNote" cols="76" rows="5" id="CategoryNote"><%=CategoryNote%></textarea></td>
            </tr>
            <tr>
                <td align="center" colspan="2" height="35" valign="bottom">
                    <input name="CategoryLoai" type="hidden" value="<%=CategoryLoai%>">
                    <input type="submit" name="Submit" value="Tạo mới">
                    <input type="button" name="Submit2" value="Đóng cửa sổ" onclick="javascript: window.close();">
                </td>
            </tr>
        </table>
    </form>
    
<!--    <script type="text/javascript">
        CKEDITOR.replace('CategoryNote');
</script>-->
</body>
</html>
<script src='../inc/news.js'></script>
<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.forms[0].CategoryNote)</script>
