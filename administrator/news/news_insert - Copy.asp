<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
 '   dim request
	'set request = new ASPupload 
	'request.codepage =65001
	Dim sError
	sError="" 'Luu ca'c thong ba'o loi
	'request.save(path)
	'categoryid=replace(request.Form("categoryid"),"'","''")
		
		'Kiem tra xem Gui tin len mot hay nhieu chuyen muc
		if trim(request.Form("categoryid"))="" then
		'Gui tin len mot chuyen muc
			'Lay gia' tri CatId cua chuyen muc
			CatId=CLng(request.form("CatId_DependRole"))
			'Kiem tra xem Chuyen muc na`y co' phu` hop voi StatusId khong
			if not IsNumeric(request.Form("StatusId")) then
				StatusId=0
			else
				StatusId=Clng((request.Form("StatusId")))
			end if
			
			
			
			if StatusId=0 then
				sError=sError & "&nbsp;-&nbsp; Trạng thái gửi tin<br>"
			else
				StatusId=trim(CheckStatusWithCategoryId(0,CatId,StatusId,Session("LstRole")))
				if len(StatusId)<>4 then
					sError=sError & "&nbsp;-&nbsp; " & StatusId & "<br>"
				end if
			end if
		else
		'Gui tin len nhieu chuyen muc
			LstCategoryId=Trim(request.form("categoryid"))
			ArrCat=Split(" " & LstCategoryId)
			CatId=ArrCat(1)
		end if
		
		'Xa'c dinh xem User co' quyen gi` trong Chuyen muc na`y
		'Editor,GroupSenior,Approver,Administrator
		strTmp=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
		if trim(request.Form("categoryid"))<>"" then
		'Lay StatusId voi truong hop ban tin va`o nhieu chuyen muc
			statusId=strTmp & strTmp
		end if
		Select case strTmp
			case "ed"
				Editor="N'" & session("user") & "'"
				GroupSenior="NULL"
				Approver="NULL"
				Administrator="NULL"
			case "se"
				Editor="NULL"
				GroupSenior="N'" & session("user") & "'"
				Approver="NULL"
				Administrator="NULL"
			case "ap"
				Editor="NULL"
				GroupSenior="NULL"
				Approver="N'" & session("user") & "'"
				Administrator="NULL"
			case "ad"
				Editor="NULL"
				GroupSenior="NULL"
				Approver="NULL"
				Administrator="N'" & session("user") & "'"
		End select

		'IsHomeNews
		if request.Form("IsHomeNews")<>"" and isnumeric(request.Form("IsHomeNews")) then
			IsHomeNews = CLng(request.Form("IsHomeNews"))
		else
			IsHomeNews = 0
		end if
		'EventId
		if isNumeric(request.form("eventid")) then
			EventId=CLng(request.form("eventid"))
		else
			EventId=0
		end if

		'languageid
		if len(request.form("languageid"))>1 then
			languageid=Trim(request.form("languageid"))
		else
			languageid="VN"
		end if
		'IsHotNews
		if request.Form("IsHotNews")<>"" and isnumeric(request.Form("IsHotNews")) then
			IsHotNews = CLng(request.Form("IsHotNews"))
		else
			IsHotNews = 0
		end if
		
		PublicationNo=0
		'Title
		if Trim(request.form("Title"))<>"" then
			Title=Trim(replace(request.form("Title"),"'","''"))
			Title=Replace(Title,"""","&quot;")
		else
			sError=sError & "&nbsp;-&nbsp; Tiêu đề<br>"
		end if
		if Trim(request.form("Tinhtrang"))<>"" then
			Tinhtrang=Trim(request.form("Tinhtrang"))
			Tinhtrang=Replace(Tinhtrang,"'","''")
			Tinhtrang=Replace(Tinhtrang,chr(13) & chr(10),"<br>")
		else
			Tinhtrang=""
		end if
		
		if Trim(request.form("BaoHanh"))<>"" then
			BaoHanh=Trim(request.form("BaoHanh"))
			BaoHanh=Replace(BaoHanh,"'","''")
			BaoHanh=Replace(BaoHanh,chr(13) & chr(10),"<br>")
		else
			BaoHanh=""
		end if		
		
			nxb=Trim(request.form("f_nxb"))
'			nxb=Replace(nxb,"'","''")
'			nxb=Replace(nxb,chr(13) & chr(10),"<br>")
			
		if Trim(request.form("f_tacgia"))<>"" then
			tacgia=Trim(request.form("f_tacgia"))
			tacgia=Replace(tacgia,"'","''")
			tacgia=Replace(tacgia,chr(13) & chr(10),"<br>")
		else
			tacgia=""
		end if	
		
		
			kichthuoc=Cstr(request.form("f_kichthuoc"))
			i = inStr(kichthuoc,"cm")
			if i = 0 then 
				i = inStr(kichthuoc,"CM")
			end if
			if i = 0 then 
				i = inStr(kichthuoc,"cM")
			end if
			if i = 0 then 
				i = inStr(kichthuoc,"Cm")
			end if
			if i= 0 then
				kichthuoc = kichthuoc+"cm"
			end if 
			
			trongluong = Trim(Request.Form("f_trongluong")) 'tungnt thêm vào ngày 28.01.2008
			namxuatban = Trim(Request.Form("f_namxuatban"))
			quy = Trim(Request.Form("f_quy"))
			
			namxuatban = Cint(namxuatban&quy)
			
		if Trim(request.form("f_giabia"))<>"" then
			giabia=Trim(request.form("f_giabia"))
			giabia=Replace(giabia,"'","''")
			giabia=Replace(giabia,chr(13) & chr(10),"<br>")
		else
			giabia=""
		end if	
			giabia=chuan_money(giabia)

		
		if request.Form("Gia")<>"" then
			Gia = Trim(request.Form("Gia"))
			Gia=Gia/100
			Gia=(giabia-giabia*Gia)
		else
			Gia = "0"
		end if
			Gia=chuan_money(Gia)

		
		if Trim(request.form("f_sotrang"))<>"" then
			sotrang=Trim(request.form("f_sotrang"))
			sotrang=Replace(sotrang,"'","''")
			sotrang=Replace(sotrang,chr(13) & chr(10),"<br>")
		else
			sotrang=""
		end if	
		
		'Description
		if Trim(request.form("Description"))<>"" then
			Description=Trim(request.form("Description"))
			Description=Replace(Description,"'","''")
			Description=Replace(Description,chr(13) & chr(10),"<br>")
		else
			sError=sError & "&nbsp;-&nbsp; Tóm tắt [Khi chọn ra trang chủ (trên) thì tóm tắt phải nhỏ hơn 260 ký tự]<br>"
		end if


		'bodyx
		if (Trim(request.form("bodyx"))<>"") then
			body=request.form("bodyx")
			body=Replace(body,"'","''")
			'Thay thế các tên miền -> để trỏ về địa chỉ tương đối.
     		body=Replace(body,"http://" & request.ServerVariables("server_name") & "/administrator/news/","")
			''body=Replace(body,"http://" & request.ServerVariables("server_name"),"")
			body=Replace(body,chr(13) & chr(10),"<br>")
		else
			sError=sError & "&nbsp;-&nbsp; Nội dung<br>"
		end if
		'smallpicturefilename
		'smallpicture =replace(request.requestedFiles("smallpicturefilename").FileName,"'","''") ' request.form("SmallPictureFileName")
		dim mt, filekey
'		mt = request.requestedFiles.keys
		
		'if (UBound(mt) <> -1) then
		'  for each filekey in request.requestedFiles.keys
		'     if filekey = "smallpicturefilename" then
		'       smallpic = smallpic + 1
			' end if  
			' if filekey="largepicturefilename" then
        '       largepic = largepic + 1
			' end if 
		'  next
		'end if
		
		'largepicturefilename
	  '  largepicture = replace(request.requestedFiles("largepicturefilename").FileName,"'","''") 'request.Files("LargePictureFileName")
		
		
		'PictureAlign
		PictureAlign=Trim(request.form("PictureAlign"))
		if PictureAlign="right" then
			PictureAlign="right"
		else
			PictureAlign="left"
		end if
		'PictureDirection
		PictureDirection=Trim(request.form("PictureDirection"))
		if not IsNumeric(PictureDirection) or PictureDirection="" then
			PictureDirection=0
		else
			PictureDirection=1
		end if
		'PictureCaption
		PictureCaption=Trim(request.form("PictureCaption"))
		PictureCaption=Replace(PictureCaption,"'","''")
		PictureCaption=Replace(PictureCaption,"""","&quot;")
		'PictureAuthor
		PictureAuthor=Trim(request.form("PictureAuthor"))
		PictureAuthor=Replace(PictureAuthor,"'","''")
		PictureAuthor=Replace(PictureAuthor,"""","&quot;")
		'Author
		Author=Trim(request.form("Author"))
		Author=Replace(Author,"'","''")
		Author=Replace(Author,"""","&quot;")
		'Source
		Source=Trim(request.form("Source"))
		Source=Replace(Source,"'","''")
		Source=Replace(Source,"""","&quot;")
		'mã sản phẩm
		if Trim(request.Form("categoryid"))<>"" then
			iCountCat	=	countCategory(CategoryID) +1
			masp = Cstr(getCategoryNote(CategoryID))&Cstr(iCountCat)
		else
			iCountCat	=	countCategory(CatId) +1
			masp = Cstr(getCategoryNote(CatId))&Cstr(iCountCat)
		end if
		
		Hethang	=Request.Form("Hethang")
		if Hethang	=	"" then
			Hethang = 0
		else
			Hethang = 1
		end if
	
		Note=Trim(request.form("Note"))
		Note=Replace(Note,"'","''")
		Note=Replace(Note,chr(13) & chr(10),"<br>")
		Note="&nbsp;&nbsp;-&nbsp;Tạo bởi: <b>" & session("user") & "</b> lúc " & Hour(now) & "h" & Minute(now) & "&quot;&nbsp;" & Day(now) & "/" & Month(now) & "/" & Year(now) & "<br>" & Note
		PictureID = 0
		if sError="" then 'Dữ liệu nhập vào hợp lệ
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			
			
			NewsId=GetMaxId("News", "NewsId", "")
			sql="insert into News (NewsId,idcode,Title,Description,Body,Author,Source,IsHomeNews"
			sql=sql & ",IsCatHomeNews,IsCatHomeNews_Below,IsHotNews,EventID,PictureID,PictureAlign,PictureDirection,LanguageID,"
			sql=sql & "PublicationNo,Creator,LastEditor,LastEditedDate,StatusId,Editor,GroupSenior,"
			sql=sql & "Approver,Administrator,Gia,giabia,Tinhtrang,Note,IsSlaveHomePageNews,tacgia,nxb,sotrang,kichthuoc"
			sql = sql & ",trongluong, namxuatban, Het"
			sql = sql & ") values ("
			sql=sql & NewsId
			sql=sql & ",N'" & masp & "'"
			sql=sql & ",N'" & Title & "'"
			sql=sql & ",N'" & Description & "'"
			sql=sql & ",N'" & body & "'"
			sql=sql & ",N'" & Author & "'"
			sql=sql & ",N'" & Source & "'"
			sql=sql & "," & IsHomeNews
			sql=sql & "," & IsCatHomeNews
			sql=sql & "," & IsCatHomeNews_Below
			sql=sql & "," & IsHotNews
			sql=sql & "," & EventID
			sql=sql & "," & PictureID
			sql=sql & ",'" & PictureAlign & "'"
			sql=sql & "," & PictureDirection
			sql=sql & ",'" & LanguageID & "'"
			sql=sql & "," & PublicationNo 
			sql=sql & ",N'" & session("user") & "'" 'Creator
			sql=sql & ",N'" & session("user") & "'" 'LastEditor
			sql=sql & ",'" & now() & "'" 'LastEditedDate
			sql=sql & ",'" & StatusId & "'" 'StatusId
			sql=sql & "," & Editor  'Editor
			sql=sql & "," & GroupSenior  'GroupSenior
			sql=sql & "," & Approver  'Approver
			sql=sql & "," & Administrator  'Administrator
			sql=sql & ",'" & Gia  &"'"
			sql=sql & ",'" & Giabia  &"'"
			sql=sql & ",N'" & Tinhtrang &"'"
			sql=sql & ",N'" & Note & "'"
			sql=sql & "," & IsSlaveHomePageNews 
			sql=sql & ",N'" & tacgia & "'"
			sql=sql & ",N'" & nxb & "'"
			sql=sql & "," & sotrang 
			sql=sql & ",N'" & kichthuoc & "'"
			sql=sql & "," & trongluong
			sql=sql & "," & namxuatban
			sql=sql & "," & Hethang
			sql = sql & ")"
			rs.open sql,con,1
			
			if trim(request.Form("categoryid"))<>"" then
				sql1=""
				for i=1 to UBound(ArrCat)
					if i<>1 then
						sql1=sql1 & ";"
					end if
					sql1=sql1 & "insert into NewsDistribution "
					sql1=sql1 & "(NewsId,CategoryId) values "
					sql1=sql1 & "(" & NewsId
					sql1=sql1 & "," & ArrCat(i) & ")"
				Next
				CatId=ArrCat(1)
			else
				sql1=sql1 & "insert into NewsDistribution "
				sql1=sql1 & "(NewsId,CategoryId) values "
				sql1=sql1 & "(" & NewsId
				sql1=sql1 & "," & CatId & ")"
			end if
			'response.write "sqlNewsDistribution=" & sql1
			'response.end()
			rs.open sql1,con,1
	'		set request.form=nothing
			set rs=nothing
			con.close
			set con=nothing
           smallPicture=Trim(request.form("SmallPictureFileName"))
		   if  smallPicture="Yes" then
            response.redirect ("news_nhapanh.asp?newsid=" & NewsId & "&catid=" & CatId & "&PictureCaption=" & PictureCaption & "&PictureAuthor=" & PictureAuthor & "&PictureAlign=" & PictureAlign & "&PictureDirection=" & PictureDirection & "&userr=" & session("user") & "&PictureId=" & 0)
			else
			'response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
			response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
		    end if
end if
	'set rs=nothing
	'con.close
	'set con=nothing
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Title_This_Page="Thêm tin & sản phẩm"
	Call header()
	
	
	
%>
  <table width="770" border="0" align="center" cellpadding="6" cellspacing="0">
   <tr>
   	<td width="20%">&nbsp;</td>
   	<td><br>
   		<font face="Arial" size="3" color="red"><strong>
   			Thiếu thông tin!
   		</strong></font><br><br>
   		<font face="Arial" size="2"><%=sError%></font><br>
   		<font face="Arial" size="2"><a href="javascript: history.go(-1);">Quay lại</a></font>
   	</td>
   </tr>
  </table>
<%Call Footer()%>
</body>
</html>
