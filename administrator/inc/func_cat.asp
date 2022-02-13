<%Session.CodePage="65001"%>
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%Sub ListTopHit(TopHit)%>
	<Select name="TopHit" id="TopHit">
		<option value="0"<%if TopHit=0 then%> selected<%End if%>>No Top Hit</option>
		<option value="1"<%if TopHit=1 then%> selected<%End if%>>Top xseo giới thiệu</option>
		<option value="2"<%if TopHit=2 then%> selected<%End if%>>Top giảm giá</option>
		<option value="3"<%if TopHit=3 then%> selected<%End if%>>Top tiêu biểu</option>
		<option value="4"<%if TopHit=4 then%> selected<%End if%>>Top bán chạy</option>
		<option value="5"<%if TopHit=5 then%> selected<%End if%>>Top nhiều người quan tâm nhất</option>
	</Select>
<%End Sub%>

<%Function GetTopHit(TopHit)
	Select case TopHit
		case 0
			GetTopHit="No Top Hit"	
		case 1
			GetTopHit="Top xseo giới thiệu"
		case 2
			GetTopHit="Top giảm giá"
		case 3
			GetTopHit="Top tiêu biểu"
		case 4
			GetTopHit="Top bán chạy"
		case 5
			GetTopHit="Top nhiều người quan tâm nhất"		
	End Select
End Function%>

<%Sub Active_luachon(Active)%>
<select name="AcTive" id="AcTive">
<option value="-1" <%if Active=-1 then %> selected <%End if%>>Khách hàng</option>
<option value="0"  <%if Active= 0 then %> selected <%End if%>>No Active</option>
<option value="1"  <%if Active= 1 then %> selected <%End if%>>Active</option>
</select> 
<%End Sub%>
<%Sub TinhTP_name(ID)
	set rsname=server.CreateObject("ADODB.Recordset")
	sqlname="SELECT * "
	sqlname=sqlname & " FROM Classified_TinhTP"
	sqlname=sqlname & " WHERE Classified_id="&ID
	rsname.open sqlname,con,1
	if  not rsname.EOF then
	Response.Write(rsname("Classified_Name"))
	End if
	rsname.Close
	set rsname=nothing
End Sub%>

<%Sub Nghenghiep_name(ID)
	set rsname=server.CreateObject("ADODB.Recordset")
	sqlname="SELECT * "
	sqlname=sqlname & " FROM nwl_nghenghiep"
	sqlname=sqlname & " WHERE nwl_id="&ID
	rsname.open sqlname,con,1
	if  not rsname.EOF then
	Response.Write(rsname("nwl_name"))
	End if
	rsname.Close
	set rsname=nothing
End Sub%>

<%Function ShowPCat_ID(CatID)
	ShowPCat_ID=0
	set rsname=server.CreateObject("ADODB.Recordset")
	sqlname="SELECT ParentCategoryId"
	sqlname=sqlname & " FROM NewsCategory"
	sqlname=sqlname & " WHERE CategoryId="&CatID
	sqlname=sqlname & " ORDER BY CategoryOrder"
	rsname.open sqlname,con,1
	if  not rsname.EOF then
	ShowPCat_ID= rsname("ParentCategoryId")
	End if
	rsname.Close
	set rsname=nothing
End Function%>

<%Function ShowCatID_Loai(CatID)
	ShowCatID_Loai=0
	set rsname=server.CreateObject("ADODB.Recordset")
	sqlname="SELECT CategoryLoai"
	sqlname=sqlname & " FROM NewsCategory"
	sqlname=sqlname & " WHERE CategoryId="&CatID
	sqlname=sqlname & " ORDER BY CategoryOrder"
	rsname.open sqlname,con,1
	if  not rsname.EOF then
	ShowCatID_Loai= rsname("CategoryLoai")
	End if
	rsname.Close
	set rsname=nothing
End Function%>

<%Sub ShowCat_name(CatID)
	set rsname=server.CreateObject("ADODB.Recordset")
	sqlname="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,ParentCategoryId"
	sqlname=sqlname & " FROM NewsCategory"
	sqlname=sqlname & " WHERE CategoryId="&CatID
	sqlname=sqlname & " ORDER BY CategoryOrder"
	rsname.open sqlname,con,1
	if  not rsname.EOF then
	Response.Write(rsname("CategoryName"))
	End if
	rsname.Close
	set rsname=nothing
End Sub%>


<%Sub ShowCat_luachon(PCatID,CatID,LanguageId)%>
<select name="Cat_name" id="Cat_name">
<option value="0">Tên Bản Tin</option>
<%
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,ParentCategoryId"
	sql=sql & " FROM NewsCategory"
	sql=sql & " WHERE languageid='" & LanguageId & "' AND ParentCategoryId="&PCatID
	sql=sql & " ORDER BY CategoryOrder"
	rs.open sql,con,1
	Do while not rs.EOF 
	Sub_PCatID=rs("CategoryId")
		set Sub_rs=server.CreateObject("ADODB.Recordset")
		Sub_sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,ParentCategoryId"
		Sub_sql=Sub_sql & " FROM NewsCategory"
		Sub_sql=Sub_sql & " WHERE languageid='" & LanguageId & "' AND ParentCategoryId="&Sub_PCatID
		Sub_sql=Sub_sql & " ORDER BY CategoryOrder"
		Sub_rs.open Sub_sql,con,1	
			Do while not Sub_rs.EOF 
				%>
					<option value="<%=Sub_rs("CategoryId")%>" <%if Sub_rs("CategoryId")=CatID then %> selected <%End if%>><%=Sub_rs("CategoryName")%></option>
				<%
			Sub_rs.MoveNext
			Loop
			Sub_rs.close
			set Sub_rs=nothing

	rs.MoveNext
	Loop
	rs.Close
	set rs=nothing
%>
</select> 
<%End Sub%>


<%Sub Update_YoungestChildren(LanguageId)
	'update youngestChildren
	sql="UPDATE newscategory "
	sql=sql & " SET youngestChildren = ( SELECT count(*) as dem "
	sql=sql & 							"FROM newscategory n2 "
	sql=sql & 							"WHERE  n2.ParentCategoryId=n1.CategoryId and "
	sql=sql & 							"		n2.LanguageId='" & LanguageId & "') "
	sql=sql & " FROM newscategory n1"
	sql=sql & " WHERE n1.LanguageId='" & LanguageId & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	set rs=nothing
End Sub%>

<%Sub Update_PrentCategoryId(LanguageId)
	Dim ArrCat(100,2)
	'Mang 2 chieu: phan tu 1 luu tru CatLevel, 2: CatId
	Dim CatParent(100)
	Dim i,j
	Dim rs1
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	i=0
	sql="SELECT CategoryId,CategoryLevel,CategoryOrder,ParentCategoryId"
	sql=sql & " FROM NewsCategory"
	sql=sql & " WHERE languageid='" & LanguageId & "'"
	sql=sql & " ORDER BY CategoryOrder"
	'Duyet Cat tu tren xuong
	rs.open sql,con,1
	Do while not rs.eof
		if i=0 or rs("CategoryLevel")=1 then
			i=1
			ArrCat(i,1)=Clng(rs("CategoryLevel"))
			ArrCat(i,2)=Clng(rs("CategoryId"))
			if Clng(rs("ParentCategoryId"))<>0 then
				sql="Update NewsCategory set ParentCategoryId=0"
				sql=sql & " WHERE CategoryId=" & rs("CategoryId")
				rs1.open sql,con,1
			end if
		else 'Neu i>0
			j=i
			Do while j>0
				if rs("CategoryLevel")>ArrCat(j,1) then
					sql="Update NewsCategory set ParentCategoryId=" & ArrCat(j,2)
					sql=sql & " WHERE CategoryId=" & rs("CategoryId")
					rs1.open sql,con,1
					Exit Do
				end if
				j=j-1
			Loop
			i=i+1
			ArrCat(i,1)=Clng(rs("CategoryLevel"))
			ArrCat(i,2)=Clng(rs("CategoryId"))
		end if
	rs.movenext
	Loop
	rs.close
	set rs1=nothing
	set rs=nothing
End sub%>

<%Sub MoveCat(moveAction,CatId,Catlevel,LanguageId)
	Dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
		Select case MoveAction
			case "left"
				if catlevel>1 then
					catlevel=catlevel-1
					sql="Update NewsCategory set CategoryLevel=" & catlevel & " where CategoryID=" & Catid
					rs.open sql,con,1
				end if		
			case "right"
				catlevel=catlevel+1
				sql="Update NewsCategory set CategoryLevel=" & catlevel & " where CategoryID=" & Catid
				rs.open sql,con,1
			case "down"
				catOrder=CatLevel
				'Tim CatOrder ngay duoi Category truyen va`o
				sql="SELECT top 1 CategoryOrder"
				sql=sql & " FROM newscategory"
				sql=sql & " WHERE LanguageId='" & LanguageId & "' and CategoryOrder>" & CatOrder
				sql=sql & " ORDER BY CategoryOrder"
				rs.open sql,con,1
				
				if rs.eof then
					rs.close
				else
					Downorder=Clng(rs("CategoryOrder"))
					rs.close
					sql="update NewsCategory set Categoryorder=" & CatOrder & "where CategoryOrder=" & Downorder
					rs.open sql,con,1
					sql="update NewsCategory set Categoryorder=" & DownOrder & "where CategoryId=" & CatId
					rs.open sql,con,1
				end if
			case "up"
				catOrder=CatLevel
				'Tim CatOrder ngay tren Category truyen va`o
				sql="SELECT top 1 CategoryOrder"
				sql=sql & " FROM newscategory"
				sql=sql & " WHERE LanguageId='" & LanguageId & "' and CategoryOrder<" & CatOrder
				sql=sql & " ORDER BY CategoryOrder desc"
				rs.open sql,con,1
				
				if rs.eof then
					rs.close
				else
					Uporder=Clng(rs("CategoryOrder"))
					rs.close
					sql="update NewsCategory set Categoryorder=" & CatOrder & "where CategoryOrder=" & Uporder
					rs.open sql,con,1
					sql="update NewsCategory set Categoryorder=" & Uporder & "where CategoryId=" & CatId
					rs.open sql,con,1
				end if
		End select
	set rs=nothing
End sub%>

<%Function GetListOfUser(Username,varOption)
	''Get list of Cat or list of Role and Category
	'varOption: False: GetListRoleOfUser
	'			True: GetListCatOfUser
	'Neu Username chua co' quyen gi` thi` tra ve chuoi rong
	sql="SELECT d.CategoryId, d.User_role,c.CategoryOrder,c.LanguageId"
	sql=sql & " FROM Userdistribution d, NewsCategory c"
	sql=sql & " WHERE d.CategoryId=c.CategoryId and username=N'" & userName & "'"
	sql=sql & " ORDER BY c.LanguageId desc,c.CategoryOrder"
	
	Dim rs,strtmp
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		sql="SELECT d.CategoryId, d.User_role"
		sql=sql & " FROM Userdistribution d"
		sql=sql & " WHERE d.username=N'" & userName & "'"
		rs.open sql,con,1
		if rs.eof then
			rs.close
			set rs=nothing
			GetListOfUser=""
		else
			GetListOfUser=rs("CategoryId")
			rs.close
			set rs=nothing
		end if
		Exit Function
	end if
	strtmp=""
	Do while not rs.eof
		if not varOption then
			strtmp=strtmp & rs("CategoryId") & rs("User_role") & " "
		else
			strtmp=strtmp & rs("CategoryId") &  " "
		end if
	rs.movenext
	Loop
	rs.close
	set rs=nothing
	GetListOfUser=strtmp
End Function%>

<%Function GetListChildrenOfListCat(LstCat)
	If lstCat="" then
		Exit Function
	end if
	
	Dim Arr,i,strTmp
	strTmp=""
	Arr=Split(" " & Trim(LstCat))
	for i=1 to UBound(Arr) 
		strTmp=strTmp & GetListChildrenCat(Arr(i)) & " "
	next
	GetListChildrenOfListCat=Trim(strTmp)
End Function%>

<%Function GetListChildrenCat(CatId)
	'Get Child CategoryId of Inpute Category.
	'Result is a string of CategoryId separated by spacebar, include Input Category
	Dim rs,i,j,Arr(50),ArrValue(100)
	Set rs=Server.CreateObject("ADODB.Recordset")
	Arr(1)=CatId
	i=1
	strCat=""
	Do 
		sql="select Categoryid from NewsCategory where ParentCategoryId=" & Arr(i)
		rs.open sql,con,1
		if not rs.eof then
			i=i-1
			Do while not rs.eof
				j=j+1
				Arrvalue(j)=rs("CategoryId")
				i=i+1
				Arr(i)=Clng(rs("CategoryId"))
			rs.movenext
			Loop
		else
			i=i-1
		end if
		rs.close
	Loop while i>0
	set rs=nothing
	Arrvalue(j+1)=CatId
	GetListChildrenCat=Trim(Join(ArrValue))
End Function%>

<%Sub ListCategoryNoOfUser(Username,LstCatUsr)
	''Get List of Category not in LstCatUsr
	sql=BuildSQL_ListCategory(LstCatUsr,"","NONE","",1,0,0)
	
	if sql="" then
		exit sub
	end if
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	
	response.Write "<select name=""categoryid"" size=""20"" multiple id=""categoryid"">" & vbNewline
		Response.Write "<option value=""0"" style=""COLOR: Blue;"">Tất cả các chuyên mục</option>" & vbNewline
	Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid")  & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red;"
			 if rs("LanguageId")="VN" then
			 	response.Write(" background-color:#FFFFFF""")
			 elseif rs("LanguageId")="EN" then
			 	response.Write(" background-color:#E6E8E9""")
			 end if
		else
			if rs("LanguageId")="VN" then
				response.Write(" style=""background-color:#FFFFFF""")
			elseif rs("LanguageId")="EN" then
				response.Write(" style=""background-color:#E6E8E9""")
			end if
		end if
		
		if Clng(rs("categoryid"))=CatSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	response.Write "</select>" & vbNewline
	
	rs.close
	set rs=nothing
End sub%>

<%Function BuildSQL_ListCategory(LstCat,LstRoleCat,LanguageId,Role,VarOption,CatStatus,ParentAndChild)
	'CatStatus =4: Tất cả các chuyên mục
	'CatStatus =0: Loại bỏ các chuyên mục có CatStatus=0
	'CatStatus =3: Loại bỏ các chuyên mục có CatStatus=0 and CatStatus=3
	
	'ParentAndChild =0 Get Only Child Category
	'				=1 Get Only Parent Category of List Input Category
	'				=2 Get Both of Parent and Child Category of List Input Category
	
	if VarOption=0 then
	'Liệt kê các chuyên mục nằm trong danh sách đưa vào
		if Role="" then
		'Liệt kê theo LstCat
			if LstCat="" then
			'Bỏ qua luôn
				BuildSQL_ListCategory=""
				Exit Function
			elseif LstCat="0" then
			'Toàn bộ các chuyên mục
				sql=GetSQL_1(Languageid,CatStatus)
			else 
			'Lấy danh sách chuyên mục trong LstCat
				sql=GetSQL_2(LstCat,LanguageId,CatStatus,ParentAndChild)
			end if 
		else'Of if Role="" then
		'Liệt kê theo LstRoleCat
			if LstRoleCat="" then
			'Bỏ qua luôn
				BuildSQL_ListCategory=""
				Exit Function
			elseif LstRoleCat="0ed" or LstRoleCat="0se" or LstRoleCat="0ap" or LstRoleCat="0ad" then
			'Coi như dữ liệu UserDistribution đã là chuẩn
				sql=GetSQL_3(LstRoleCat,LanguageId,Role,CatStatus)
			else
			'Lay danh sa'ch chuyen muc trong LstCat
				sql=GetSQL_4(LstRoleCat,LanguageId,Role,CatStatus)
			end if 'Of if LstRoleCat="" then
		end if 
	else
	'Liệt kê các chuyên mục không nằm trong danh sách đưa vào
		if Role="" then
		'Liệt kê theo LstCat
			if LstCat="0" then
			'Bỏ qua luôn vì không còn chuyên mục nào kô nằm trong danh sách đưa vào
				BuildSQL_ListCategory=""
				Exit Function
			elseif LstCat="" then
			'Lấy toàn bộ các chuyên mục
				sql=GetSQL_1(Languageid,CatStatus)
			else 
			'Lấy danh sách chuyên mục trong LstCat
				sql=GetSQL_5(LstCat,LanguageId,CatStatus)
			end if 
		else'Of if Role="" then
		'Liệt kê theo LstRoleCat
			'Không có trường hợp nào nên tắt function này, he he
		end if 
	end if
	
	BuildSQL_ListCategory=sql
End Function%>

<%Sub ListCategoryOfUser(Username,LstCatUsr)
	'Get List of Cat In LstCatUsr
	response.Write"<Script language=""JavaScript"">" & vbNewline &_
	"function fRolesRemoveSubmit(thisvalue)" & vbNewline &_
	"{" & vbNewline &_
		"document.fRoles.cat.value=thisvalue;" & vbNewline &_
		"document.fRoles.action.value=""Remove"";" & vbNewline &_
		"document.fRoles.submit();" & vbNewline &_
	"}" & vbNewline &_
	"</Script>" & vbNewline
	
	IF LstCatUsr="" THEN
	'Danh sách chuyên mục truyền vào không chứa chuyên mục nào
		Exit Sub
	ELSEIF LstCatUsr="0" THEN
		'Là toàn bộ chuyên mục
		Response.Write "<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""2"">" & vbNewline &_
			"	<tr>" & vbNewline &_
			"		<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
			"			&nbsp;&#8226;&nbsp;T&#7845;t c&#7843; c&#225;c chuy&#234;n m&#7909;c</font></td>" & vbNewline &_
			"		<td align=""center""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
			"			" & GetRoles(Username,0) & "</font></td>" & vbNewline &_
			"		<td align=""center""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
			"		<a href=""javascript: fRolesRemoveSubmit(0);"">" & vbNewline &_
			"			<img src=""../images/delete.jpg"" width=""13"" height=""16"" border=""0"" title=""Xóa bỏ"">" & vbNewline &_
			"		</a></font></td>" & vbNewline &_
			"	</tr>" & vbNewline &_
			"</table>" & vbNewline
	ELSEIF Trim(LstCatUsr)<>CStr(GiaoLuuCategoryId) then
	'Nếu quyền khác giao lưu thì
		'response.write LstCatUsr & "=" & GiaoLuuCategoryId
		'response.end
		sql=BuildSQL_ListCategory(LstCatUsr,"","NONE","",0,0,2)

		Dim rs
		set rs=server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbNewline

		boo=false
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
			for i=2 to Clng(rs("CategoryLevel"))
				response.Write("&nbsp;&nbsp;")
			Next 
			if Clng(rs("YoungestChildren"))=0 and Clng(rs("CategoryLevel"))<>1 then
				response.write "&nbsp;-&nbsp;" & rs("CategoryName") & "<br>"
			else
				if Clng(rs("CategoryLevel"))=1 then
					response.write "&nbsp;&#8226;&nbsp;<b>" & rs("CategoryName") & "</b><br>"
				else
					response.write "&nbsp;&#8226;&nbsp;" & rs("CategoryName") & "<br>"
				end if
			end if
		    Response.Write	"</font></td>" & vbNewline
			strTmp=GetRoles(Username,rs("CategoryId"))
			if  strTmp<>"" then
			    Response.Write	"<td align=""center"" valign=""middle""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
					strTmp & "</font></td>" & vbNewline &_ 
					"<td align=""center"" valign=""middle"">" & vbNewline &_
					"<a href=""javascript: fRolesRemoveSubmit(" & rs("CategoryId") & ");"">" & vbnewline &_
					"	<img src=""../images/delete.jpg"" width=""13"" height=""16"" border=""0"" title=""Xóa bỏ"">" & vbNewline &_
					"</a></td>" & vbNewline
			else
				response.Write("<td></td><td></td>")
			end if
   			Response.Write"</tr>" & vbNewline
		rs.movenext
		Loop
		Response.Write("</table>")
		rs.close
		set rs=nothing
	ELSE
	''Nếu là phỏng vấn trực tuyến
		
	END IF
End sub%>

<%Function GetListChildrenRole_FromListRole(LstRole)
	Dim strTmp,ArrTmp1,ArrCat(100),ArrRole(100),strCat
	Dim i,j,k
	
	ArrTmp1=Split(" " & Trim(LstRole))
	strTmp=""
	strCat=""
	k=0

	For i=1 to UBound(ArrTmp1)
		CategoryRole=Mid(ArrTmp1(i),len(ArrTmp1(i))-1,2)
		'Get Role of Category
		CategoryId=Mid(ArrTmp1(i),1,len(ArrTmp1(i))-2)
		'Get CatId of Category
		strTmp=GetListChildrenCat(CategoryId)
		ArrTmp2=Split(" " & Trim(strTmp))

		for j=1 to UBound(ArrTmp2) 
			k=k+1
			ArrCat(k)=ArrTmp2(j)
			ArrRole(k)=CategoryRole
		next
	Next
	strTmp=""
	for i=1 to k 
		strTmp=strTmp & " " &  ArrCat(i) & ArrRole(i)
	next
	GetListChildrenRole_FromListRole=Trim(strTmp)
End Function%>
<%Function GetRoleOfCat_FromListRole(CatId,LstRole)
	Select case LstRole
		case "0ed"
			GetRoleOfCat_FromListRole="ed"
			Exit Function
		case "0se"
			GetRoleOfCat_FromListRole="se"
			Exit Function
		case "0ap"
			GetRoleOfCat_FromListRole="ap"
			Exit Function
		case "0ad"
			GetRoleOfCat_FromListRole="ad"
			Exit Function
	End Select
	Dim Cat_ed,Cat_se,Cat_ap,Cat_ad,sLstRole
	Cat_ed=CatId & "ed"
	Cat_se=CatId & "se"
	Cat_ap=CatId & "ap"
	Cat_ad=CatId & "ad"
	sLstRole=GetListChildrenRole_FromListRole(LstRole)

	if Instr(" " & sLstRole & " "," " & Cat_ed & " ")>0 then
		GetRoleOfCat_FromListRole="ed"
	elseif Instr(" " & sLstRole & " "," " & Cat_se & " ")>0 then
		GetRoleOfCat_FromListRole="se"
	elseif Instr(" " & sLstRole & " "," " & Cat_ap & " ")>0 then
		GetRoleOfCat_FromListRole="ap"
	elseif Instr(" " & sLstRole & " "," " & Cat_ad & " ")>0 then
		GetRoleOfCat_FromListRole="ad"
	end if
End Function%>

<%Function GetRoles(Username,CatId)
	''Get name of Role
	sql="SELECT user_role FROM userdistribution where Username=N'" & Username & "' and CategoryId=" & CatId
	Dim rs1
	set rs1=server.CreateObject("ADODB.Recordset")
	rs1.open sql,con,1
	if rs1.eof then
		GetRoles=""
		rs1.close
		set rs1=nothing
		Exit Function
	end if
	Select case rs1("user_role")
		case "ed"
			GetRoles="B.T&#7853;p Vi&#234;n"
		case "se"
			GetRoles="Hi&#7879;u &#272;&#237;nh"
		case "ap"
			GetRoles="Ph&#7909; Tr&#225;ch"
		case "ad"
			GetRoles="Qu&#7843;n Tr&#7883;"
	End Select
	rs1.close
	set rs1=nothing
End Function%>

<%Function GetListParentOfListCat(LstCat)
	If lstCat="" then
		Exit Function
	end if
	
	Dim Arr,i,strTmp
	strTmp=""
	Arr=Split(" " & Trim(LstCat))
	for i=1 to UBound(Arr) 
		strTmp=strTmp & GetListParentCat(Arr(i)) & " "
	next
	GetListParentOfListCat=Trim(strTmp)
End Function%>

<%Function GetListParentCat(CatId)
	'Get Tree List CategoryId of Inpute Category.
	'Result is a string of CategoryId separated by spacebar, not include Input Category
	Dim i,ArrValue(100)
	i=0
	Dim rs1
	set rs1=Server.CreateObject("ADODB.Recordset")

	  PCatId=CatId
	  Do while PCatId<>0
		sql_GetListParentCat="select ParentCategoryId from NewsCategory where CategoryId='" & PCatId &"'"
		rs1.open sql_GetListParentCat,con,1
			PCatId=Clng(rs1("ParentCategoryId"))
			if PcatId<>0 then
				i=i+1
				ArrValue(i)=rs1("ParentCategoryId")
			end if
		rs1.close
	  Loop
	GetListParentCat=Trim(Join(ArrValue))
End Function%>

<%Function GetListParentCatNameOfCatId(CatId)
	''Get Tree List Name Of CategoryId of Inpute Category.
	'Result is a string of CategoryId's separated by spacebar, include Input Category
	Dim i,strArrValue
	i=0
	Dim rs1
	set rs1=Server.CreateObject("ADODB.Recordset")

	  PCatId=CatId
	  Do while PCatId<>0
	  	i=i+1
		sql_GetListParentCat="select ParentCategoryId,CategoryName from NewsCategory where CategoryId=" & PCatId
		rs1.open sql_GetListParentCat,con,1
			PCatId=Clng(rs1("ParentCategoryId"))
			if i=1 then
				strArrValue=rs1("CategoryName")
			else
				strArrValue=rs1("CategoryName") & " -> " & strArrValue
			end if
		rs1.close
	  Loop
	GetListParentCatNameOfCatId=strArrValue
End Function%>

<%Sub ListTreeCategory(CatId)
	''Display Tree Category adding javascript
	response.Write 	"<script src=""/administrator/js/lket_script1.js""></script>" & vbNewline &_
					"<script src=""/administrator/js/lket_script2.js""></script>" & vbNewline

	sql="SELECT CategoryName,CategoryId,CategoryOrder,LanguageId, CategoryLevel, YoungestChildren, ParentCategoryId"
	sql=sql & " FROM newscategory"
	sql=sql & " WHERE CategoryStatus=1 or CategoryStatus=2 or CategoryStatus=0 or CategoryStatus=3"
	sql=sql & " and TopHit = 0 ORDER BY languageId desc, CategoryOrder"
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		exit sub
	end if
	
	Dim i, pos, ArrCat,strLstParentId, strTmp
	strLstParentId=GetListParentCat(CatId)
	
	Response.Write 	"<Script language=""JavaScript"">" & vbNewline &_
					"foldersTree = gFld(""<font style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none\"">Danh s&#225;ch chuy&#234;n m&#7909;c</font>"");" & vbNewline &_
					"foldersTree.treeID = ""checkboxTree"";" & vbNewline
	strTmp=""
	i=0
	Do while not rs.eof
		i=i+1
		if rs("CategoryLevel")=1 then
			Response.Write	"Cat" & rs("CategoryId") & "=insFld(foldersTree, gFld(""<a href=\""" & Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId") & "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.write " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		elseif rs("YoungestChildren")<>0 then
			Response.Write "Cat" & rs("CategoryId") & "= insFld(Cat" & rs("ParentCategoryId") & ", gFld(""<a href=\""" & Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId") & "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.Write " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		else'Youngest Children=0
			Response.Write "insDoc(Cat" & rs("ParentCategoryId") & ", gLnk(""<a href=\""" & Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId") & "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.Write  " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		end if
		
		if CompareRoleCat(strLstParentId,rs("CategoryId"))<>0 then
			strTmp=strTmp & "clickOnNode(" & i & ");"
		end if
	rs.movenext
	Loop
	rs.close
	set rs=nothing
	Response.Write("</Script>" & vbNewline)
	Response.Write	"<script language=""JavaScript"">initializeDocument()</script>" & vbNewline 
	Response.Write	"<script>" & strTmp & "</script>" & vbNewline
End sub%>

<%Function CompareRole(Role1,Role2)
	'Reuturn Which Role greater or wicker
	Dim Arr(4)
	Arr(1)="ed"
	Arr(2)="se"
	Arr(3)="ap"
	Arr(4)="ad"
	for i=1 to 4 
		if Arr(i)=lcase(Role1) then
			Exit For
		end if
	next
	for j=1 to 4 
		if Arr(j)=lcase(Role2) then
			Exit For
		end if
	next
	CompareRole=i-j
End Function%>

<%Function GetSQL_1(LanguageId,CatStatus)
	boo=false
	sql="SELECT *"
	sql=sql & " FROM newsCategory WHERE TopHit = 0 "
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & "  and CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " and LanguageId='" & LanguageId & "'"
		end if
	end if
	sql=sql & "  ORDER BY LanguageId desc, CategoryOrder"
	GetSQL_1=sql
End Function%>

<%Function GetSQL_2(LstCat,LanguageId,CatStatus,ParentAndChild)
	Dim strTmp,ArrCat2
	strTmp=Trim(GetListChildrenOfListCat(LstCat)) & " " & Trim(GetListParentOfListCat(LstCat))
	ArrCat2=Split(" " & Trim(strTmp))

	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,YoungestChildren,LanguageId"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	for i=1 to UBound(ArrCat2)
		if i=1 then
			if boo then
				sql=sql & " and (CategoryId=" & ArrCat2(i)
			else
				sql=sql & " WHERE (CategoryId=" & ArrCat2(i)
			end if
		else
			if ArrCat2(i)<>"" then 
				sql=sql & " or CategoryId=" & ArrCat2(i)
			end if
		end if
	Next
	sql=sql & ")"
	sql=sql & " ORDER BY LanguageId desc,CategoryOrder"
	'response.write sql
	'response.end
	GetSQL_2=sql
End Function%>

<%Function GetSQL_3(LstRoleCat,LanguageId,Role,CatStatus)
	SELECT CASE lcase(LstRoleCat)
		CASE "0ed"
			if CompareRole("ed",Role)>=0 then
				sql=GetSQL_1(LanguageId,CatStatus)
			else
				sql=""
			end if
		CASE "0se"
			if CompareRole("se",Role)>=0 then
				sql=GetSQL_1(LanguageId,CatStatus)
			else
				sql=""
			end if
		CASE "0ap"
			if CompareRole("ap",Role)>=0 then
				sql=GetSQL_1(LanguageId,CatStatus)
			else
				sql=""
			end if
		CASE "0ad"
			sql=GetSQL_1(LanguageId,CatStatus)
	END SELECT
	GetSQL_3=sql
End Function%>

<%Function GetSQL_4(LstRoleCat,LanguageId,Role,CatStatus)
	Dim strTmp,ArrTmp1,ArrCat(100),ArrRole(100)
	Dim i,j,k
	
	ArrTmp1=Split(" " & Trim(LstRoleCat))
	strTmp=""
	
	k=0

	For i=1 to UBound(ArrTmp1)
		CategoryRole=Mid(ArrTmp1(i),len(ArrTmp1(i))-1,2)
		'Get Role of Category
		CategoryId=Mid(ArrTmp1(i),1,len(ArrTmp1(i))-2)
		'Get CatId of Category
		strTmp=GetListChildrenCat(CategoryId)
		ArrTmp2=Split(" " & Trim(strTmp))

		for j=1 to UBound(ArrTmp2) 
			k=k+1
			ArrCat(k)=ArrTmp2(j)
			ArrRole(k)=CategoryRole
		next
	Next
	'K saved the UBound value of ArrCat and ArrRole
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,YoungestChildren,LanguageId"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	j=0
	for i=1 to k
		if CompareRole(ArrRole(i),Role)>=0 then
			j=j+1
			if j=1 then
				if boo then
					sql=sql & " and (CategoryId=" & ArrCat(i)
				else
					sql=sql & " WHERE (CategoryId=" & ArrCat(i)
				end if
			else
				sql=sql & " or CategoryId=" & ArrCat(i)
			end if
		end if
	Next
	if j>0 then
		sql=sql & ")"
	else
		'Khong co' quyen va`o chuyen muc na`o, tra ve rong
		sql=""
		GetSQL_4=sql
		exit Function
	end if
	sql=sql & " ORDER BY LanguageId desc,CategoryOrder"

	GetSQL_4=sql	
End Function%>
<%Function GetSQL_5(LstCat,LanguageId,CatStatus)
	Dim strTmp,ArrCat1,ArrCat2
	ArrCat1=Split(" " & Trim(LstCat))
	strTmp=""
	For i=1 to UBound(ArrCat1)
		strTmp=strTmp & GetListChildrenCat(ArrCat1(i)) & " "
	Next
	strTmp=Trim(strTmp)
	ArrCat2=Split(" " & strTmp)
	
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,YoungestChildren,LanguageId"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	for i=1 to UBound(ArrCat2)
		if i=1 then
			if boo then
				sql=sql & " and CategoryId<>" & ArrCat2(i)
			else
				sql=sql & " WHERE CategoryId<>" & ArrCat2(i)
			end if
		else
			sql=sql & " and CategoryId<>" & ArrCat2(i)
		end if
	Next
	sql=sql & " ORDER BY LanguageId desc,CategoryOrder"
	
	GetSQL_5=sql
End Function%>

<%Sub List_Category(CatSelect,CatTitle,LanguageId,CategoryLoai)
	'CatSelect=CategoryID is choosen.
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="SELECT *"
	sql=sql & " FROM NewsCategory"
	if LanguageId="NONE" then
		if CategoryLoai<>-1 then
			sql=sql + " CategoryLoai = '"& CategoryLoai &"'"
		end if	
	else
		sql=sql & " Where Languageid='" & LanguageId & "'"
		if CategoryLoai<>-1 then
			sql=sql + " and CategoryLoai = '"& CategoryLoai &"'"
		end if		
	end if

	sql=sql & " ORDER BY CategoryOrder"
	rs.Open sql, con, 1
	
	response.Write "<select name=""categoryid"" id=""categoryid"">" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid")  & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red; background-color:#FFFFFF"""
		end if
		
		if Cint(rs("categoryid"))=Cint(CatSelect) then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>


<%Sub List_Category_Name(CatSelect,NameSelect,CatTitle,LanguageId,CategoryLoai)
    
	'CatSelect=CategoryID is choosen.
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="SELECT *"
	sql=sql & " FROM NewsCategory"
	if LanguageId="NONE" then
		if CategoryLoai<>-1 then
			sql=sql + " CategoryLoai = '"& CategoryLoai &"'"
		end if	
	else
		sql=sql & " Where Languageid='" & LanguageId & "'"
		if CategoryLoai<>-1 then
			sql=sql + " and CategoryLoai = '"& CategoryLoai &"'"
		end if		
	end if

	sql=sql & " ORDER BY CategoryOrder"
	rs.Open sql, con, 1
	
	response.Write "<select name="&NameSelect&" id="&NameSelect&" class='form-control' style='width:200px;'>" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid")  & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red; background-color:#FFFFFF"""
		end if
		
		if Cint(rs("categoryid"))=Cint(CatSelect) then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>

<%Sub List_CategoryWithoutSelectTag(CatSelect)
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="SELECT *"
	sql=sql & " FROM NewsCategory"
	sql=sql & " where TopHit=0 ORDER BY LanguageId desc, CategoryOrder"
	rs.Open sql, con, 1
	
    Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid") & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red; background-color:#FFFFFF"""
		end if
		
		if Clng(rs("categoryid"))=CatSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
End Sub%>

<%Sub List_Category_Depend_Role(CatSelect, CatTitle, LanguageId, LstRole, Role, CatStatus,CategoryLoai)
    if CategoryLoai <> 0 then CategoryLoai = "disabled"
LstRole= "0ad"
Role="ap"
CatStatus=0
IF LstRole="0ed" or LstRole="0se" or LstRole="0ap" or LstRole="0ad" THEN
	sql=GetSQL_3(LstRole,LanguageId,Role,CatStatus)
ELSE
	Dim strTmp,ArrTmp1,ArrCat(100),ArrRole(100),strCat
	Dim i,j,k
	
	ArrTmp1=Split(" " & Trim(LstRole))
	strTmp=""
	strCat=""
	k=0
	
	For i=1 to UBound(ArrTmp1)
		CategoryRole=Mid(ArrTmp1(i),len(ArrTmp1(i))-1,2)
		'Get Role of Category
		CategoryId=Mid(ArrTmp1(i),1,len(ArrTmp1(i))-2)
		'Get CatId of Category
		'response.write "CategoryRole=" & CategoryRole & ",CategoryId=" & CategoryId
		strTmp=GetListChildrenCat(CategoryId)
		
		ArrTmp2=Split(" " & Trim(strTmp))
		
		for j=1 to UBound(ArrTmp2) 
			k=k+1
			ArrCat(k)=ArrTmp2(j)
			ArrRole(k)=CategoryRole
		next
	Next

	'K saved the UBound value of ArrCat and ArrRole
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,YoungestChildren,LanguageId"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & " "
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0  and CategoryStatus<>3 "
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	j=0
	
	for i=1 to k
		if CompareRole(ArrRole(i),Role)>=0 and IsNumeric(ArrCat(i)) then
			j=j+1
			if j=1 then
				if boo then
					sql=sql & " and (CategoryId=" & ArrCat(i)
				else
					sql=sql & " WHERE (CategoryId=" & ArrCat(i)
				end if
			else
				sql=sql & " or CategoryId=" & ArrCat(i)
			end if
			strCat=strCat & " " & ArrCat(i)
		end if
	Next
	
	if j>0 then
		'Co' quyen va`o mot so chuyen muc
		strParentCat=Trim(GetListParentOfListCat(Trim(strCat)))
		if strParentCat<>"" then 
			ArrTmp1=Split(" " & Trim(strParentCat))
			'Neu co' chuyen muc cha
			for i=1 to UBound(ArrTmp1)
			  if IsNumeric(ArrTmp1(i)) then
				sql=sql & " or CategoryId=" & ArrTmp1(i)
			  end if
			Next
		end if
		sql=sql & ")"
		sql=sql & " and TopHit = 0 ORDER BY LanguageId desc,CategoryOrder"
	else
		'Khong co' quyen va`o chuyen muc na`o, tra ve rong
		sql=""
		Exit Sub
	end if
END IF 
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
	"	function checkme(thisoption,checkvalue)" & vbNewline &_
	"	{" & vbNewline &_
	"		if (checkvalue==0)" & vbNewline &_
	"		{" & vbNewline &_
	"			alert (""Bạn không được chọn chuyên mục này!"");" & vbNewline &_
	"			thisoption.options[0].selected=true" & vbNewline &_
	"		}" & vbNewline &_
	"	}" & vbNewline &_
	"</script>" & vbNewline
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	'response.Write("SQL=" & sql)
	rs.open sql,con,1
	response.Write "<select name=""CatId_DependRole"" id=""CatId_DependRole"" onChange=""javascript: checkme(this,this.value);"" "&CategoryLoai&">" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option "
		if CompareRoleCat(strParentCat,rs("CategoryId"))>0 and CompareRoleCat(strCat,rs("CategoryId"))=0 then
		'Nam trong strParentcat nhung ko nam trong strCat
			response.Write "value=""0""  style=""COLOR: Red; background-color:#FFFFF0"""
		else
			response.Write "value=""" & rs("CategoryId") & """"
			if rs("CategoryLevel")=1 then
			 	response.Write " style=""COLOR: Blue; background-color:#FFFFFF"""
			end if
		end if

		if Clng(rs("categoryid"))=CatSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>

<%Sub ListTreeCategory_WithRole(CatId, CatTitle, LanguageId, LstRole, Role, CatStatus)
IF LstRole="0ed" or LstRole="0se" or LstRole="0ap" or LstRole="0ad" THEN
	sql=GetSQL_3(LstRole,LanguageId,Role,CatStatus)
ELSE
	Dim strTmp,ArrTmp1,ArrCat(100),ArrRole(100),strCat
	Dim i,j,k
	
	ArrTmp1=Split(" " & Trim(LstRole))
	strTmp=""
	strCat=""
	k=0

	For i=1 to UBound(ArrTmp1)
		CategoryRole=Mid(ArrTmp1(i),len(ArrTmp1(i))-1,2)
		'Get Role of Category
		CategoryId=Mid(ArrTmp1(i),1,len(ArrTmp1(i))-2)
		'Get CatId of Category
		strTmp=GetListChildrenCat(CategoryId)
		ArrTmp2=Split(" " & Trim(strTmp))

		for j=1 to UBound(ArrTmp2) 
			k=k+1
			ArrCat(k)=ArrTmp2(j)
			ArrRole(k)=CategoryRole
		next
	Next

	'K saved the UBound value of ArrCat and ArrRole
	sql="SELECT *"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	j=0
	for i=1 to k
		if CompareRole(ArrRole(i),Role)>=0 then
			j=j+1
			if j=1 then
				if boo then
					sql=sql & " and (CategoryId=" & ArrCat(i)
				else
					sql=sql & " WHERE (CategoryId=" & ArrCat(i)
				end if
			else
				sql=sql & " or CategoryId=" & ArrCat(i)
			end if
			strCat=strCat & " " & ArrCat(i)
		end if
	Next

	if j>0 then
		'Co' quyen va`o mot so chuyen muc
		strParentCat=Trim(GetListParentOfListCat(Trim(strCat)))
		if strParentCat<>"" then 
			ArrTmp1=Split(" " & Trim(strParentCat))
			for i=1 to UBound(ArrTmp1)
			  if IsNumeric(ArrTmp1(i)) then
				 sql=sql & " or CategoryId=" & ArrTmp1(i)
			  end if
			Next
		end if
		sql=sql & ")"
		sql=sql & " ORDER BY LanguageId desc,CategoryOrder"
	else
		'Khong co' quyen va`o chuyen muc na`o, tra ve rong
		sql=""
		Exit Sub
	end if
END IF 
	''response.write sql
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		exit sub
	end if

	Dim pos, strLstParentId
	strLstParentId=GetListParentCat(CatId)
	response.Write 	"<script src=""/administrator/js/lket_script1.js""></script>" & vbNewline &_
					"<script src=""/administrator/js/lket_script2.js""></script>" & vbNewline
	Response.Write 	"<Script language=""JavaScript"">" & vbNewline &_
					"foldersTree = gFld(""<font style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none\"">" & CatTitle & "</font>"");" & vbNewline &_
					"foldersTree.treeID = ""checkboxTree"";" & vbNewline
	strTmp=""
	i=0
	
	Do while not rs.eof
		i=i+1
		if rs("CategoryLevel")=1 then
			Response.Write	"Cat" & rs("CategoryId") & "=insFld(foldersTree, gFld(""<a href=\"""
			
			if CompareRoleCat(strParentCat,rs("CategoryId"))>0 and CompareRoleCat(strCat,rs("CategoryId"))=0 then
			'Đoạn code này dùng để loại bỏ các đường link ParentCatId (Chỉ hiển thị ra để làm minh họa)
			'Nam trong strParentcat nhung ko nam trong strCat
				Response.write "#"
			else
				Response.write Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId")
			end if
			Response.write "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.write " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		elseif rs("YoungestChildren")<>0 then
			Response.Write "Cat" & rs("CategoryId") & "= insFld(Cat" & rs("ParentCategoryId") & ", gFld(""<a href=\"""

			if CompareRoleCat(strParentCat,rs("CategoryId"))>0 and CompareRoleCat(strCat,rs("CategoryId"))=0 then
			'Đoạn code này dùng để loại bỏ các đường link ParentCatId (Chỉ hiển thị ra để làm minh họa)
			'Nam trong strParentcat nhung ko nam trong strCat
				Response.write "#"
			else
				Response.write Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId")
			end if
			Response.write "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.Write " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		else'Youngest Children=0
			Response.Write "insDoc(Cat" & rs("ParentCategoryId") & ", gLnk(""<a href=\""" & Request.ServerVariables("SCRIPT_NAME") & "?catid=" & rs("CategoryId") & "\"" style=\""COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 8.5pt; FONT-STYLE: normal;"
			if Clng(rs("CategoryId"))=CatId then
				Response.Write " FONT-WEIGHT: bold;"
			else
				Response.Write " FONT-WEIGHT: normal;"
			end if
			Response.Write  " TEXT-DECORATION: none\"">" & rs("CategoryName") & "</a>""));" & vbNewline
		end if
		
		if CompareRoleCat(strLstParentId,rs("CategoryId"))<>0 then
			strTmp=strTmp & "clickOnNode(" & i & ");"
		end if
	rs.movenext
	Loop
	rs.close
	set rs=nothing
	Response.Write("</Script>" & vbNewline)
	Response.Write	"<script language=""JavaScript"">initializeDocument()</script>" & vbNewline 
	Response.Write	"<script>" & strTmp & "</script>" & vbNewline
End sub%>

<%Sub List_Category_MultiChoose_Depend_Role(CatSelect, CatTitle,LanguageId,LstRole,Role,CatStatus)
IF LstRole="0ed" or LstRole="0se" or LstRole="0ap" or LstRole="0ad" THEN
	sql=GetSQL_3(LstRole,LanguageId,Role,CatStatus)
ELSE
	Dim strTmp,ArrTmp1,ArrCat(100),ArrRole(100),strCat
	Dim i,j,k
	
	ArrTmp1=Split(" " & Trim(LstRole))
	strTmp=""
	strCat=""
	k=0

	For i=1 to UBound(ArrTmp1)
		CategoryRole=Mid(ArrTmp1(i),len(ArrTmp1(i))-1,2)
		'Get Role of Category
		CategoryId=Mid(ArrTmp1(i),1,len(ArrTmp1(i))-2)
		'Get CatId of Category
		strTmp=GetListChildrenCat(CategoryId)
		ArrTmp2=Split(" " & Trim(strTmp))

		for j=1 to UBound(ArrTmp2) 
			k=k+1
			ArrCat(k)=ArrTmp2(j)
			ArrRole(k)=CategoryRole
		next
	Next

	'K saved the UBound value of ArrCat and ArrRole
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,YoungestChildren,LanguageId"
	sql=sql & " From NewsCategory"
	boo=false
	if CatStatus=0 then
		sql=sql & ""
		boo=true
	elseif CatStatus=3 then
		sql=sql & " WHERE CategoryStatus<>0 and CategoryStatus<>3"
		boo=true
	end if
	if LanguageId<>"NONE" then
		if boo then
			sql=sql & " and LanguageId='" & LanguageId & "'"
		else
			sql=sql & " WHERE LanguageId='" & LanguageId & "'"
		end if
		boo=true
	end if
	j=0
	for i=1 to k
		if CompareRole(ArrRole(i),Role)>=0 then
			j=j+1
			if j=1 then
				if boo then
					sql=sql & " and (CategoryId=" & ArrCat(i)
				else
					sql=sql & " WHERE (CategoryId=" & ArrCat(i)
				end if
			else
				sql=sql & " or CategoryId=" & ArrCat(i)
			end if
			strCat=strCat & " " & ArrCat(i)
		end if
	Next
	if j>0 then
		'Co' quyen va`o mot so chuyen muc
		strParentCat=Trim(GetListParentOfListCat(Trim(strCat)))
		if strParentCat<>"" then 
			ArrTmp1=Split(" " & Trim(strParentCat))
		end if
		for i=1 to UBound(ArrTmp1)
		  if IsNumeric(ArrTmp1(i)) then
			 sql=sql & " or CategoryId=" & ArrTmp1(i)
		  end if
		Next
		sql=sql & ")"
		sql=sql & " ORDER BY LanguageId desc,CategoryOrder"
	else
		'Khong co' quyen va`o chuyen muc na`o, tra ve rong
		sql=""
		Exit Sub
	end if
END IF 
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
	"	function checkme(thisoption,checkvalue)" & vbNewline &_
	"	{" & vbNewline &_
	"		if (checkvalue==0)" & vbNewline &_
	"		{" & vbNewline &_
	"			alert (""Bạn không được chọn chuyên mục này!"");" & vbNewline &_
	"			for(i=0;i<thisoption.length;i++)" & vbNewline &_
	"			{" & vbNewline &_
	"				if ((thisoption.options[i].value==0) && (thisoption.options[i].selected=true))" & vbNewline &_
	"					thisoption.options[i].selected=false;" & vbNewline &_
	"			}" & vbNewline &_
	"		}" & vbNewline &_
	"	}" & vbNewline &_
	"</script>" & vbNewline
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	'response.Write("SQL=" & sql)
	rs.open sql,con,1
	response.Write "<select name=""CatId_DependRole"" id=""CatId_DependRole"" onChange=""javascript: checkme(this,this.value);"" size=""15"" multiple>" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option "
		if CompareRoleCat(strParentCat,rs("CategoryId"))>0 and CompareRoleCat(strCat,rs("CategoryId"))=0 then
		'Nam trong strParentcat nhung ko nam trong strCat
			response.Write "value=""0"" style=""COLOR: Red; background-color:#FFFFF0"""
		else
			response.Write "value=""" & rs("CategoryId") & """"
			if rs("CategoryLevel")=1 then
			 	response.Write " style=""COLOR: Blue; background-color:#FFFFFF"""
			end if
		end if

		if Clng(rs("categoryid"))=CatSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>

<%Function GetSQL_For_Search(LstCat,LstRoleCat,username,SearchFilter)
	'Nếu SearchFilter="Marked" Chỉ liệt kê các tin được đánh dấu
	'Nếu SearchFilter="Waiting" Chỉ liệt kê các tin chờ duyệt
	'Nếu SearchFilter="Edit" Chỉ liệt kê các tin được cấp trên yêu cầu sửa lại
	'Else Liệt kê tất các tin trong phạm vi quyền
	
	Dim strTmp
	strTmp=""
	strLstCat=LstCat
	strLstRole=LstRoleCat
	strUserName=username
	SELECT CASE strLstRole
	case "0ed" 'Biên tập viên
		'Chỉ có 2 quyền
		strTmp="("
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=edma và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='edma' and n.Editor=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của hiệu đính StatusId=seed và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='seed' and n.Editor=N'" & strUserName & "')"
			case "Waiting"
				'Do Biên tập viên không có quyền duyệt tin nên khi Biên tập viên chọn tìm kiếm theo tiêu chí này
				'thì để khóa tìm kiếm là CategoryId=0 -> Tức là luôn không tìm thấy tin nào.
				strTmp=strTmp & "(d.CategoryId=0)"
			case else 'Liệt kê cả 2 trường hợp
				'Quyền xem tin đánh dấu của mình StatusId=edma và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='edma' and n.Editor=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của hiệu đính StatusId=seed và Editor=strUserName
				strTmp=strTmp & " or (n.StatusId='seed' and n.Editor=N'" & strUserName & "')"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")"
	case "0se" 'Hiệu đính
		'Có 3 quyền
		strTmp="(" 'Mở ngoặc
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=sema và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='sema' and n.GroupSenior=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của phụ trách StatusId=apse và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='apse' and n.GroupSenior=N'" & strUserName & "')"
			case "Waiting"
				'Quyền duyệt tin do Biên tập viên thuộc quyền phụ trách gửi lên StatusId=edse
				strTmp=strTmp & "(n.StatusId='edse')"
			case else
				'Quyền xem tin đánh dấu của mình StatusId=sema và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='sema' and n.GroupSenior=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của phụ trách StatusId=apse và GroupSenior=strUserName
				strTmp=strTmp & " or (n.StatusId='apse' and n.GroupSenior=N'" & strUserName & "')"
				'Quyền duyệt tin do Biên tập viên thuộc quyền phụ trách gửi lên StatusId=edse
				strTmp=strTmp & " or (n.StatusId='edse')"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")" 'Đóng ngoặc
	case "0ap" 'Phụ trách
		'Có 4 quyền
		strTmp="(" 'Mở ngoặc
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=apma và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='apma' and n.Approver=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của tổng phụ trách StatusId=adap và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='adap' and n.Approver=N'" & strUserName & "')"
			case "Waiting"
				'Quyền duyệt tin do Hiệu đính thuộc quyền phụ trách gửi lên StatusId=seap
				strTmp=strTmp & "(n.StatusId='seap')"
			case else
				'Quyền xem tin đánh dấu của mình StatusId=apma và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='apma' and n.Approver=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của tổng phụ trách StatusId=adap và Approver=strUserName
				strTmp=strTmp & " or (n.StatusId='adap' and n.Approver=N'" & strUserName & "')"
				'Quyền duyệt tin do Hiệu đính thuộc quyền phụ trách gửi lên StatusId=seap
				strTmp=strTmp & " or (n.StatusId='seap')"
				'Quyền Sửa lại các tin đã gửi lên mạng
				strTmp=strTmp & " or (n.StatusId='apap' or n.StatusId='adad')"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")" 'Đóng ngoặc
	case "0ad" 'Tổng phụ trách
		'Có 3 quyền
		strTmp="(" 'Mở ngoặc
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu mình StatusId=adma và Administrator=strUserName
				strTmp=strTmp & "(n.StatusId='adma' and n.Administrator=N'" & strUserName & "')"
			case "Edit"
				'Do Tổng phụ trách không nằm dưới quyền của ai nên nên khi chọn tìm kiếm theo tiêu chí này
				'thì để khóa tìm kiếm là CategoryId=0 -> Tức là luôn không tìm thấy tin nào.
				strTmp=strTmp & "(d.CategoryId=0)"
			case "Waiting"
				'Quyền duyệt tin do phụ trách gửi lên StatusId=apad
				strTmp=strTmp & "(n.StatusId='apad')"
			case else
				'Quyền xem tin đánh dấu mình StatusId=adma và Administrator=strUserName
				strTmp=strTmp & "(n.StatusId='adma' and n.Administrator=N'" & strUserName & "')"
				'Quyền duyệt tin do phụ trách gửi lên StatusId=apad
				strTmp=strTmp & " or (n.StatusId='apad')"
				'Quyền Sửa lại các tin đã gửi lên mạng
				strTmp=strTmp & " or (n.StatusId='apap' or n.StatusId='adad')"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")" 'Đóng ngoặc
	case else 'Quyền pha tạp Chắc chắn không có quyền Administrator, nên không tính.
		Dim ArrLstCat,ArrED,ArrAP,ArrSE
		Dim strED,strSE,strAP
		strED=""
		strSE=""
		strAP=""
		
		strLstCat=GetListChildrenOfListCat(strLstCat)
		ArrLstCat=Split(" " & strLstCat)
		strLstRole=GetListChildrenRole_FromListRole(strLstRole)

		for i=1 to UBound(ArrLstCat)
			if Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "ed ")>0 then
				if strED<>"" then
					strED=strED & " or "
				end if
				strED=strED & "d.CategoryId=" & ArrLstCat(i)
			elseif Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "se ")>0 then
				if strSE<>"" then
					strSE=strSE & " or "
				end if
				strSE=strSE & "d.CategoryId=" & ArrLstCat(i)
			elseif Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "ap ")>0 then
				if strAP<>"" then
					strAP=strAP & " or "
				end if
				strAP=strAP & "d.CategoryId=" & ArrLstCat(i)
			end if
		next
		
		if strED<>"" then
			'Nhóm strED làm 1 cụm. VD: (d.CategoryId=3 or d.CategoryId=6)
			strED="(" & strED & ")"
			'Chỉ có 2 quyền
			strTmp="("
			Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=edma và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='edma' and n.Editor=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của hiệu đính StatusId=seed và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='seed' and n.Editor=N'" & strUserName & "')"
			case "Waiting"
				'Do Biên tập viên không có quyền duyệt tin nên khi Biên tập viên chọn tìm kiếm theo tiêu chí này
				'thì để khóa tìm kiếm là CategoryId=0 -> Tức là luôn không tìm thấy tin nào.
				strTmp=strTmp & "(d.CategoryId=0)"
			case else 'Liệt kê cả 2 trường hợp
				'Quyền xem tin đánh dấu của mình StatusId=edma và Editor=strUserName
				strTmp=strTmp & "(n.StatusId='edma' and n.Editor=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của hiệu đính StatusId=seed và Editor=strUserName
				strTmp=strTmp & " or (n.StatusId='seed' and n.Editor=N'" & strUserName & "')"
			end select	
			'Vì ở đây chỉ có quyền Editor tại một số các chuyên mục nên:
			strTmp=strTmp & ")" & " and " & strED
			'Nhóm strTmp lại
			strED="(" & strTmp & ")"
		end if

		if strSE<>"" then
			'Nhóm strSE làm 1 cụm. VD: (d.CategoryId=3 or d.CategoryId=6)
			strSE="(" & strSE & ")"
			'Có 3 quyền
			strTmp="(" 'Mở ngoặc
			
			Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=sema và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='sema' and n.GroupSenior=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của phụ trách StatusId=apse và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='apse' and n.GroupSenior=N'" & strUserName & "')"
			case "Waiting"
				'Quyền duyệt tin do Biên tập viên thuộc quyền phụ trách gửi lên StatusId=edse
				strTmp=strTmp & "(n.StatusId='edse')"
			case else
				'Quyền xem tin đánh dấu của mình StatusId=sema và GroupSenior=strUserName
				strTmp=strTmp & "(n.StatusId='sema' and n.GroupSenior=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của phụ trách StatusId=apse và GroupSenior=strUserName
				strTmp=strTmp & " or (n.StatusId='apse' and n.GroupSenior=N'" & strUserName & "')"
				'Quyền duyệt tin do Biên tập viên thuộc quyền phụ trách gửi lên StatusId=edse
				strTmp=strTmp & " or (n.StatusId='edse')"
			end select	
			'Vì ở đây chỉ có quyền Hiệu đính tại một số các chuyên mục nên:
			strTmp=strTmp & ")" & " and " & strSE
			'Nhóm strTmp lại
			strTmp="(" & strTmp & ")"
			if strED<>"" then
				strSE=strTmp & " or " & strED
			else
				strSE=strTmp
			end if
		else
			strSE=strED
		end if
		'Response.write strAP & ":" & SearchFilter
		if strAP<>"" then
			'Nhóm strAP làm 1 cụm. VD: (d.CategoryId=3 or d.CategoryId=6)
			strAP="(" & strAP & ")"
			'Có 4 quyền
			strTmp="(" 'Mở ngoặc
			
			Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu của mình StatusId=apma và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='apma' and n.Approver=N'" & strUserName & "')"
			case "Edit"
				'Quyền sửa tin theo yêu cầu của tổng phụ trách StatusId=adap và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='adap' and n.Approver=N'" & strUserName & "')"
			case "Waiting"
				'Quyền duyệt tin do Hiệu đính thuộc quyền phụ trách gửi lên StatusId=seap
				strTmp=strTmp & "(n.StatusId='seap')"
			case else
				'Quyền xem tin đánh dấu của mình StatusId=apma và Approver=strUserName
				strTmp=strTmp & "(n.StatusId='apma' and n.Approver=N'" & strUserName & "')"
				'Quyền sửa tin theo yêu cầu của tổng phụ trách StatusId=adap và Approver=strUserName
				strTmp=strTmp & " or (n.StatusId='adap' and n.Approver=N'" & strUserName & "')"
				'Quyền duyệt tin do Hiệu đính thuộc quyền phụ trách gửi lên StatusId=seap
				strTmp=strTmp & " or (n.StatusId='seap')"
				'Quyền Sửa lại các tin đã gửi lên mạng
				strTmp=strTmp & " or (n.StatusId='apap' or n.StatusId='adad')"
			end select
			
			'Vì ở đây chỉ có quyền Hiệu đính tại một số các chuyên mục nên:
			strTmp=strTmp & ")" & " and " & strAP

			'Nhóm strTmp lại
			strTmp="(" & strTmp & ")"
			if strSE<>"" then
				strAP=strTmp & " or " & strSE
			else
				strAP=strTmp
			end if
		else
			strAP=strSE
		end if
		strTmp="(" & strAP & ")"
	END SELECT
	GetSQL_For_Search=strTmp
End Function%>

<%Function GetSQL_For_Search_2Role(LstCat,LstRoleCat,username,SearchFilter)
	'Có chức năng xây dựng câu lệnh query sql như hàm GetSQL_For_Search. Khác nhau:
	'	- Quy quyền của User về 2 quyền: BTV=HĐ (nhóm 1), PT=TPT (nhóm 2)
	'	- Trạng thái của tin chỉ có 5 trạng thái:
	'		+> new : "Mới" do độc giả gửi lên (cả nhóm 1x2 đều có quyền xem hoặc sửa)
	'		+> mama: "Đánh dấu" do bất kỳ User nào đánh dấu (ai đánh dấu thì người đó xem hoặc sửa được)
	'		+> eded: "Chờ duyệt" đã được nhóm 1 duyệt gửi lên nhóm 2 (cả nhóm 1x2 đều có quyền xem hoặc sửa)
	'		+> apap: "Lên mạng" (chỉ nhóm 2 mới có quyền xem hoặc sửa)
	'		+> adad: "Lên mạng" (chỉ nhóm 2 mới có quyền xem hoặc sửa)
	
	'SearchFilter	="Marked" Chỉ liệt kê các tin,bài được đánh dấu
	'				="Waiting" Chỉ liệt kê các tin,bài chờ duyệt
	'				="New" Chỉ liệt kê các tin,bài mới
	'				Else Liệt kê tất các tin trong phạm vi quyền
	
	Dim strTmp
	strTmp=""
	strLstCat=LstCat
	strLstRole=LstRoleCat
	strUserName=username
	SELECT CASE strLstRole
	case "0ed", "0se" 'Biên tập viên
		'Chỉ có 3 quyền
		strTmp="("
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
			case "New"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & "c.StatusId='new'"
			case "Waiting"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & "c.StatusId='eded'"
			case else 'Liệt kê cả 3 trường hợp
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & " or c.StatusId='new'"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & " or c.StatusId='eded'"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")"

	case "0ap", "0ad" 'Phụ trách
		'Có 4 quyền
		strTmp="(" 'Mở ngoặc
		Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
			case "New"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & "c.StatusId='new'"
			case "Waiting"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & "c.StatusId='eded'"
			case else 'Liệt kê cả 3 trường hợp
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & " or c.StatusId='new'"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & " or c.StatusId='eded'"
				'Quyền xem và sửa tin lên mạng
				strTmp=strTmp & " or c.StatusId='apap' or c.StatusId='adad'"
		end select
		'Vì ở đây là tất cả các chuyên mục nên không quan tâm đến chuyên mục nào
		strTmp=strTmp & ")" 'Đóng ngoặc
	case else 'Quyền pha tạp Chắc chắn không có quyền Administrator, nên không tính.
		Dim ArrLstCat,ArrED,ArrAP,ArrSE
		Dim strED,strSE,strAP
		strED=""
		strAP=""
		
		strLstCat=GetListChildrenOfListCat(strLstCat)
		ArrLstCat=Split(" " & strLstCat)
		strLstRole=GetListChildrenRole_FromListRole(strLstRole)

		for i=1 to UBound(ArrLstCat)
			if 	Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "ed ")>0 or Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "se ")>0 then
				if strED<>"" then
					strED=strED & " or "
				end if
				strED=strED & "d.CategoryId=" & ArrLstCat(i)
			elseif Instr(" " & strLstRole & " "," " & ArrLstCat(i) & "ap ")>0 then
				if strAP<>"" then
					strAP=strAP & " or "
				end if
				strAP=strAP & "d.CategoryId=" & ArrLstCat(i)
			end if
		next
		
		if strED<>"" then
			'Nhóm strED làm 1 cụm. VD: (d.CategoryId=3 or d.CategoryId=6)
			strED="(" & strED & ")"
			'Chỉ có 3 quyền
			strTmp="("
			Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
			case "New"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & "c.StatusId='new'"
			case "Waiting"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & "c.StatusId='eded'"
			case else 'Liệt kê cả 3 trường hợp
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & " or c.StatusId='new'"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & " or c.StatusId='eded'"
			end select
			'Vì ở đây chỉ có quyền Editor tại một số các chuyên mục nên:
			strTmp=strTmp & ")" & " and " & strED
			'Nhóm strTmp lại
			strED="(" & strTmp & ")"
		end if
		
		if strAP<>"" then
			'Nhóm strAP làm 1 cụm. VD: (d.CategoryId=3 or d.CategoryId=6)
			strAP="(" & strAP & ")"
			'Có 4 quyền
			strTmp="(" 'Mở ngoặc
			Select case SearchFilter
			case "Marked"
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
			case "New"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & "c.StatusId='new'"
			case "Waiting"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & "c.StatusId='eded'"
			case else 'Liệt kê cả 3 trường hợp
				'Quyền xem tin đánh dấu 
				strTmp=strTmp & "(c.StatusId='mama' and c.LastEditor=N'" & strUserName & "')"
				'Quyền xem và sửa tin mới
				strTmp=strTmp & " or c.StatusId='new'"
				'Quyền xem và sửa lại tin chờ duyệt
				strTmp=strTmp & " or c.StatusId='eded'"
				'Quyền xem và sửa tin lên mạng
				strTmp=strTmp & " or c.StatusId='apap' or c.StatusId='adad'"
			end select
			'Vì ở đây chỉ có quyền phụ trách tại một số các chuyên mục nên:
			strTmp=strTmp & ")" & " and " & strAP
			'Nhóm strTmp lại
			strTmp="(" & strTmp & ")"
			if strED<>"" then
				strAP=strTmp & " or " & strED
			else
				strAP=strTmp
			end if
		else
			strAP=strED
		end if
		strTmp="(" & strAP & ")"
	END SELECT
	GetSQL_For_Search_2Role=strTmp
End Function%>
<%Function CheckCatInLstCat(LstCat,CatId)
	if Instr(" " & LstCat & " "," " & CatId & " ")>0 then
		CheckCatInLstCat=true
	else
		CheckCatInLstCat=false
	end if
End Function%>

<%''''''''The thao''''''''''''%>
<%Sub MoveLoaigiai(moveAction,CatId,Catlevel,LanguageId)
	Dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
		Select case MoveAction
			case "left"
				if catlevel>1 then
					catlevel=catlevel-1
					sql="Update TT_LOAIGIAI set CategoryLevel=" & catlevel & " where CategoryID=" & Catid
					rs.open sql,con,1
				end if		
			case "right"
				catlevel=catlevel+1
				sql="Update TT_LOAIGIAI set CategoryLevel=" & catlevel & " where CategoryID=" & Catid
				rs.open sql,con,1
			case "down"
				catOrder=CatLevel
				'Tim CatOrder ngay duoi Category truyen va`o
				sql="SELECT top 1 CategoryOrder"
				sql=sql & " FROM TT_LOAIGIAI"
				sql=sql & " WHERE LanguageId='" & LanguageId & "' and CategoryOrder>" & CatOrder
				sql=sql & " ORDER BY CategoryOrder"
				rs.open sql,con,1
				
				if rs.eof then
					rs.close
				else
					Downorder=Clng(rs("CategoryOrder"))
					rs.close
					sql="update TT_LOAIGIAI set Categoryorder=" & CatOrder & "where CategoryOrder=" & Downorder
					rs.open sql,con,1
					sql="update TT_LOAIGIAI set Categoryorder=" & DownOrder & "where CategoryId=" & CatId
					rs.open sql,con,1
				end if
			case "up"
				catOrder=CatLevel
				'Tim CatOrder ngay tren Category truyen va`o
				sql="SELECT top 1 CategoryOrder"
				sql=sql & " FROM TT_LOAIGIAI"
				sql=sql & " WHERE LanguageId='" & LanguageId & "' and CategoryOrder<" & CatOrder
				sql=sql & " ORDER BY CategoryOrder desc"
				rs.open sql,con,1
				
				if rs.eof then
					rs.close
				else
					Uporder=Clng(rs("CategoryOrder"))
					rs.close
					sql="update TT_LOAIGIAI set Categoryorder=" & CatOrder & "where CategoryOrder=" & Uporder
					rs.open sql,con,1
					sql="update TT_LOAIGIAI set Categoryorder=" & Uporder & "where CategoryId=" & CatId
					rs.open sql,con,1
				end if
		End select
	set rs=nothing
End sub%>


<%Sub Update_PrentCategoryId_Loaigiai(LanguageId)
	Dim ArrCat(100,2)
	'Mang 2 chieu: phan tu 1 luu tru CatLevel, 2: CatId
	Dim CatParent(100)
	Dim i,j
	Dim rs1
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	i=0
	sql="SELECT CategoryId,CategoryLevel,CategoryOrder,ParentCategoryId"
	sql=sql & " FROM TT_LOAIGIAI"
	sql=sql & " WHERE languageid='" & LanguageId & "'"
	sql=sql & " ORDER BY CategoryOrder"
	'Duyet Cat tu tren xuong
	rs.open sql,con,1
	Do while not rs.eof
		if i=0 or rs("CategoryLevel")=1 then
			i=1
			ArrCat(i,1)=Clng(rs("CategoryLevel"))
			ArrCat(i,2)=Clng(rs("CategoryId"))
			if Clng(rs("ParentCategoryId"))<>0 then
				sql="Update TT_LOAIGIAI set ParentCategoryId=0"
				sql=sql & " WHERE CategoryId=" & rs("CategoryId")
				rs1.open sql,con,1
			end if
		else 'Neu i>0
			j=i
			Do while j>0
				if rs("CategoryLevel")>ArrCat(j,1) then
					sql="Update TT_LOAIGIAI set ParentCategoryId=" & ArrCat(j,2)
					sql=sql & " WHERE CategoryId=" & rs("CategoryId")
					rs1.open sql,con,1
					Exit Do
				end if
				j=j-1
			Loop
			i=i+1
			ArrCat(i,1)=Clng(rs("CategoryLevel"))
			ArrCat(i,2)=Clng(rs("CategoryId"))
		end if
	rs.movenext
	Loop
	rs.close
	set rs1=nothing
	set rs=nothing
End sub%>

<%Sub Update_YoungestChildren_Loaigiai(LanguageId)
	'update youngestChildren
	sql="UPDATE TT_LOAIGIAI "
	sql=sql & " SET youngestChildren = ( SELECT count(*) as dem "
	sql=sql & 							"FROM TT_LOAIGIAI n2 "
	sql=sql & 							"WHERE  n2.ParentCategoryId=n1.CategoryId and "
	sql=sql & 							"		n2.LanguageId='" & LanguageId & "') "
	sql=sql & " FROM TT_LOAIGIAI n1"
	sql=sql & " WHERE n1.LanguageId='" & LanguageId & "'"
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	set rs=nothing
End Sub%>

<%Sub List_Category_Loaigiai(CatSelect, CatTitle,LanguageId)
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT *"
	sql=sql & " FROM TT_LOAIGIAI"
	if LanguageId="NONE" then
	else
		sql=sql & " Where Languageid='" & LanguageId & "'"
	end if
	sql=sql & " ORDER BY CategoryOrder"
	rs.Open sql, con, 1
	
	response.Write "<select name=""categoryid"" id=""categoryid"">" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid")  & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red; background-color:#FFFFFF"""
		end if
		
		if Clng(rs("categoryid"))=CatSelect then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>

<%Sub ShowTT_LoaiGiai(CatID,loai)%>
<select name="Loaigiai" id="Loaigiai">
<option value="0">Tên Giải</option>
<%
''Loai=0 chi Loai giai cha
''Loai=1 chi tat ca cac loai giai
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT CategoryId,CategoryName,CategoryLevel,CategoryOrder,ParentCategoryId"
	sql=sql & " FROM TT_LoaiGiai"
if Loai=0 then	sql=sql & " WHERE ParentCategoryId=0"
	sql=sql & " ORDER BY CategoryOrder"
	rs.open sql,con,1
	Do while not rs.EOF 
		%>
		<option value="<%=rs("CategoryId")%>" <%if rs("CategoryId")=CatID then %> selected <%End if%>><%=rs("CategoryName")%></option>
		<%
	rs.MoveNext
	Loop
	rs.Close
	set rs=nothing
%>
</select> 
<%End Sub%>

<%Sub ShowTT_LoaiGiai_Name(ID)%>
<%
''Loai=0 chi Loai giai cha
''Loai=1 chi tat ca cac loai giai
	set rsLoaiGiai_Name=server.CreateObject("ADODB.Recordset")
	sqlLoaiGiai_Name="SELECT CategoryId,CategoryName"
	sqlLoaiGiai_Name=sqlLoaiGiai_Name & " FROM TT_LoaiGiai"
	sqlLoaiGiai_Name=sqlLoaiGiai_Name & " WHERE CategoryId="&ID
	rsLoaiGiai_Name.open sqlLoaiGiai_Name,con,1
	if not rsLoaiGiai_Name.EOF  then Response.Write(rsLoaiGiai_Name("CategoryName"))
	rsLoaiGiai_Name.Close
	set rsLoaiGiai_Name=nothing
%>
<%End Sub%>

<%
Function getCategoryNote(CategoryID)
	set rsGetCat=server.CreateObject("ADODB.Recordset")
	sqlGetCat="SELECT CategoryId,CategoryNote"
	sqlGetCat=sqlGetCat & " FROM NewsCategory"
	sqlGetCat=sqlGetCat & " WHERE CategoryId='"&CategoryID&"'"
	rsGetCat.open sqlGetCat,con,1
	if not rsGetCat.eof then
		getCategoryNote	=	rsGetCat("CategoryNote")
	else
		getCategoryNote	=	"MSP"
	end if
	set rsGetCat = nothing
end Function
function countCategory(CategoryID)
	set rsCountCat=server.CreateObject("ADODB.Recordset")
	sqlCountCat="SELECT COUNT(NewsID) as iTotalID"
	sqlCountCat=sqlCountCat & " FROM V_News"
	sqlCountCat=sqlCountCat & " WHERE CategoryId='"&CategoryID&"'"
	rsCountCat.open sqlCountCat,con,1
	countCategory	=	rsCountCat("iTotalID")
	set rsCountCat = nothing
end function
%>

<%
function isCategoryNote(Note)
	set rsisCat=server.CreateObject("ADODB.Recordset")
	sqlisCat="SELECT CategoryId,CategoryNote"
	sqlisCat=sqlisCat & " FROM NewsCategory"
	sqlisCat=sqlisCat & " WHERE CategoryNote=N'"&Note&"'"
	rsisCat.open sqlisCat,con,1
	if not rsisCat.eof then
		isCategoryNote	=	True
	else
		isCategoryNote	=	False
	end if
	set rsisCat = nothing
end function


%>


<%''''''''''End  the thao'''''''''%>