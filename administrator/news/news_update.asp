<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->

<%
	'Kiem tra du lieu nhap vao
	'Dim request 'Su dung Asprequest
	Dim sError
	sError="" 'Luu ca'c thong ba'o loi
	'Set request = Server.CreateObject("Persits.request")

	'request.SetMaxSize 1000000, True 'Dat kich co request la` 1MB
	'Upload.codepage=65001
    '	request.Save
		
		'NewsId
		if request.Form("NewsId")<>"" and isnumeric(request.Form("NewsId")) then
			NewsId = CLng(request.Form("NewsId"))
		else
			NewsId = 0
		end if
		'sCatId
		if request.Form("sCatId")<>"" and isnumeric(request.Form("sCatId")) then
			sCatId = CLng(request.Form("sCatId"))
		else
			sCatId = 0
		end if
		
		'Kiem tra xem Gui tin len mot hay nhieu chuyen muc
		if trim(request.Form("categoryid"))="" then
		'Trước gửi tin lên 1 chuyên mục bây giờ vẫn gửi lên một chuyên mục
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
				StatusId=trim(CheckStatusWithCategoryId(NewsId,CatId,StatusId,Session("LstRole")))
				if len(StatusId)<>4 then
					sError=sError & "&nbsp;-&nbsp; " & StatusId & "<br>"
				end if
			end if
		elseif sCatId<>0 then
		'Trước gửi lên nhiều chuyên mục bây giờ gửi vào một chuyên mục
		'Do chọn trạng thái gửi tin là Đánh dấu, Gửi trở lại cấp dưới, hoặc Gửi lên cấp trên
		'Với Button Submit là: Gửi tin với trạng thái lựa chọn
			'Kiem tra xem Chuyen muc na`y co' phu` hop voi StatusId khong
			if not IsNumeric(request.Form("StatusId")) then
				StatusId=0
			else
				StatusId=Clng((request.Form("StatusId")))
			end if
			CatId=sCatId
			if StatusId=0 then
				sError=sError & "&nbsp;-&nbsp; Trạng thái gửi tin<br>"
			elseif StatusId=4 then
				sCatId=0
			else
				StatusId=trim(CheckStatusWithCategoryId(NewsId,CatId,StatusId,Session("LstRole")))
				if len(StatusId)<>4 then
					sError=sError & "&nbsp;-&nbsp; " & StatusId & "<br>"
				end if
			end if
		end if
		
		if trim(request.Form("categoryid"))<>"" and sCatId=0 then
		'Gui tin len nhieu chuyen muc hoặc
		'Submit Button=Gửi tin với trạng thái lựa chọn 
		'và lựa chọn = Gửi lên mạng thì giữ nguyên danh sách Categoryid
			LstCategoryId=Trim(request.form("categoryid"))
			ArrCat=Split(" " & LstCategoryId)
			CatId=ArrCat(1)
		end if
		
		'Xa'c dinh xem User co' quyen gi` trong Chuyen muc na`y
		'Editor,GroupSenior,Approver,Administrator
		strTmp=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
		if trim(request.Form("categoryid"))<>"" and sCatId=0 then
		'Lay StatusId voi truong hop ban tin va`o nhieu chuyen muc
		'Submit Button=Gửi tin với trạng thái lựa chọn 
		'và lựa chọn = Gửi lên mạng thì giữ nguyên danh sách Categoryid
			statusId=strTmp & strTmp
		end if
		Select case strTmp
			case "ed"
				Editor=session("user")
				GroupSenior=""
				Approver=""
				Administrator=""
			case "se"
				Editor=""
				GroupSenior=session("user")
				Approver=""
				Administrator=""
			case "ap"
				Editor=""
				GroupSenior=""
				Approver=session("user")
				Administrator=""
			case "ad"
				Editor=""
				GroupSenior=""
				Approver=""
				Administrator=session("user")
		End select
		'old_PictureId
		if request.Form("old_PictureId")<>"" and isnumeric(request.Form("old_PictureId")) then
			old_PictureId = CLng(request.Form("old_PictureId"))
		else
			old_PictureId = 0
		end if
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
		'IsCatHomeNews
		if request.Form("IsCatHomeNews")<>"" and isnumeric(request.Form("IsCatHomeNews")) then
			IsCatHomeNews = CLng(request.Form("IsCatHomeNews"))
		else
			IsCatHomeNews = 0
		end if
		'IsCatHomeNews_Below
		if request.Form("IsCatHomeNews_Below")<>"" and isnumeric(request.Form("IsCatHomeNews_Below")) then
			IsCatHomeNews_Below = CLng(request.Form("IsCatHomeNews_Below"))
		else
			IsCatHomeNews_Below = 0
		end if
		masp=Request.Form("f_masp")
		'IsSlaveHomePageNews
		if request.Form("IsSlaveHomePageNews")<>"" and isnumeric(request.Form("IsSlaveHomePageNews")) then
			IsSlaveHomePageNews = CLng(request.Form("IsSlaveHomePageNews"))
		else
			IsSlaveHomePageNews = 0
		end if
		
		if request.Form("f_giabia")<> "" then
			Giabia = Chuan_money(request.Form("f_giabia"))
		else
			Giabia=0
		end if	

		if request.Form("CK")<> "" then
			Gia = Cint(request.Form("CK"))
			Gia=Gia/100
			Gia=(giabia-giabia*Gia)
		else
			sError=sError & "Đề nghị nhập gía sách<br>"
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

		'PublicationNo
		'Or có một công thức để tính số báo.
		'if request.Form("PublicationNo")<>"" then
		'	if isnumeric(request.Form("PublicationNo")) then
		'		PublicationNo=CLng(request.Form("PublicationNo"))
		'	else
		'		sError=sError & "&nbsp;-&nbsp; Số báo<br>"
		'	end if
		'else
		'	sError=sError & "&nbsp;-&nbsp; Số báo<br>"
		'end if
		PublicationNo=0
		'SubTitle
		SubTitle=Trim(replace(request.form("SubTitle"),"'","''"))
		SubTitle=Replace(SubTitle,"""","&quot;")
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
			sError=sError & "Thiếu tình trạng của sách<br>"
		end if
		
		if Trim(request.form("f_nxb"))<>"" then
			nxb=Trim(request.form("f_nxb"))
			nxb=Replace(nxb,"'","''")
		else
			nxb=""
		end if
'		Response.Write(nxb) 'Anh Tuân thêm dòng này vào làm gì thế? (tungnt, 15/02/2008)
'		Response.End()
		
		if Trim(request.form("f_tacgia"))<>"" then
			tacgia=Trim(request.form("f_tacgia"))
			tacgia=Replace(tacgia,"'","''")
		else
			sError=sError & "Đề nghị nhập tác giả<br>"
		end if	
		
		if Trim(request.form("f_kichthuoc"))<>"" then
			kichthuoc=Trim(request.form("f_kichthuoc"))
			kichthuoc=Replace(kichthuoc,"'","''")
			
		else
			sError=sError & "Đề nghị nhập kích thước theo chuẩn(VD:12x20 cm), dạng thập phân phải dc làm tròn<br>"
		end if	
		
		if Trim(request.form("f_sotrang"))<>"" then
			sotrang=Trim(request.form("f_sotrang"))
			sotrang=Replace(sotrang,"'","''")
			sotrang=Replace(sotrang,chr(13) & chr(10),"<br>")
		else
			sError=sError & "Đề nghị nhập số trang<br>"
		end if	
		
		'Description
		if (Trim(request.form("Description"))<>"" AND IsCatHomeNews=0) OR (Len(Trim(request.form("Description"))) < 1000 AND IsCatHomeNews=1) then
			Description=Trim(request.form("Description"))
			Description=Replace(Description,"'","''")
			Description=Replace(Description,chr(13) & chr(10),"<br>")
		else
			sError=sError & "&nbsp;-&nbsp; Tóm tắt [Khi chọn ra trang chủ (trên) thì tóm tắt phải nhỏ hơn 250 ký tự không dấu]<br>"
		end if
		'bodyx
		if (Trim(request.form("bodyx"))<>"") then
			body=Trim(request.form("bodyx"))
			body=Replace(body,"'","''")
			'Thay thế các tên miền -> để trỏ về địa chỉ tương đối.
			body=Replace(body,"http://" & Request.ServerVariables("server_name") & "/administrator/news/","")
			'body=Replace(body,"http://" & Request.ServerVariables("server_name"),"")
			
			body=Replace(body,"&shy;","")
			body=Replace(body,chr(13) & chr(10),"<br>")
		else
			sError=sError & "&nbsp;-&nbsp; Nội dung<br>"
		end if
		'RemoveImage
		if request.Form("RemoveImage")<>"" and isnumeric(request.Form("RemoveImage")) then
			RemoveImage = CLng(request.Form("RemoveImage"))
		else
			RemoveImage = 0
		end if
		'RemoveLargeImage
		if request.Form("RemoveLargeImage")<>"" and isnumeric(request.Form("RemoveLargeImage")) then
			RemoveLargeImage = CLng(request.Form("RemoveLargeImage"))
		else
			RemoveLargeImage = 0
		end if
		'smallpicturefilename
		
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
		'old_Note
		old_Note=Trim(request.form("old_Note"))
		old_Note=Replace(old_Note,"'","''")
		old_Note=Replace(old_Note,chr(13) & chr(10),"<br>")
		'Note
		Note=Trim(request.form("Note"))
		Note=Replace(Note,"'","''")
		Note=Replace(Note,chr(13) & chr(10),"<br>")
		Note=old_Note & "&nbsp;&nbsp;-&nbsp;Sửa bởi: <b>" & session("user") & "</b> lúc " & Hour(now) & "h" & Minute(now) & """&nbsp;" & Day(now) & "/" & Month(now) & "/" & Year(now) & "<br>" & Note
	    smallPicture=Trim(request.form("SmallPictureFileName"))
		
		Hethang	=Cint(Request.Form("Hethang"))
'		if Hethang	=	"" then
'			Hethang = 0
'		else
'			Hethang = 1
'		end if
		
					
		trongluong = Trim(Request.Form("f_trongluong")) 'tungnt thêm vào ngày 28.01.2008

		namxuatban = Trim(Request.Form("f_namxuatban"))
		quy = Trim(Request.Form("f_quy"))
		
		namxuatban = Cint(namxuatban&quy)
		
		if sError="" then 'Dữ liệu nhập vào hợp lệ
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			if old_PictureId=0 then
			'Trước chưa có ảnh
				'if smallPicture="Yes" then				
					PictureId=0
				'else
					'PictureId=0
				'end if
			else
			'Trước có ảnh
				if RemoveImage=1 then
				'Bây giờ xóa ảnh
					PictureId=0
				else
				'Bây giờ Update Ảnh
					PictureId=old_PictureId
					sql="Update Picture set "
					sql=sql & "PictureCaption=N'" & PictureCaption & "'"
					
					
					sql=sql & ",PictureAuthor=N'" & PictureAuthor & "'"
					sql=sql & ",CategoryID=" & CatId
					sql=sql & " where PictureId=" & PictureId
					rs.open sql,con,1
				end if
			end if
			sql="Update News set"
			sql=sql & " idcode = N'" & masp & "'"
			sql=sql & " ,Title=N'" & Title & "'"
			sql=sql & ",Description=N'" & Description & "'"
			sql=sql & ",body=N'" & body & "'"
			sql=sql & ",Author=N'" & Author & "'"
			sql=sql & ",Source=N'" & Source & "'"
			sql=sql & ",IsHomeNews=" & IsHomeNews
			sql=sql & ",IsCatHomeNews=" & IsCatHomeNews
			sql=sql & ",IsCatHomeNews_Below=" & IsCatHomeNews_Below
			sql=sql & ",IsHotNews=" & IsHotNews
			sql=sql & ",IsSlaveHomePageNews=" & IsSlaveHomePageNews
			sql=sql & ",Gia='"&Gia&"'"	
			sql=sql & ",Giabia='"&Giabia&"'"			
			sql=sql & ",Tinhtrang=N'"& Tinhtrang & "'"
			sql=sql & ",EventID=" & EventID
			sql=sql & ",PictureID=" & PictureID
			sql=sql & ",PictureAlign='" & PictureAlign & "'"
			sql=sql & ",PictureDirection=" & PictureDirection
			sql=sql & ",LanguageID='" & LanguageID & "'"
			sql=sql & ",PublicationNo=" & PublicationNo 
			sql=sql & ",LastEditor=N'" & userr & "'" 'LastEditor
			sql=sql & ",LastEditedDate='" & now() & "'" 'LastEditedDate
			sql=sql & ",StatusId='" & StatusId & "'" 'StatusId
			if Editor<>"" then
				sql=sql & ",Editor='" & Editor & "'"
			end if
			if GroupSenior<>"" then
				sql=sql & ",GroupSenior='" & GroupSenior & "'"
			end if
			if Approver<>"" then
				sql=sql & ",Approver='" & Approver & "'"
			end if
			if Administrator<>"" then
				sql=sql & ",Administrator='" & Administrator & "'"
			end if
			sql=sql & ",Note=N'" & Note & "'"
	
			sql=sql & ",tacgia=N'"&tacgia&"'"	
			sql=sql & ",nxb=N'"&nxb&"'"	
			sql=sql & ",sotrang='"&sotrang&"'"	
			sql=sql & ",kichthuoc='"&kichthuoc&"'"
			sql=sql & ",trongluong='"&trongluong&"'"
			sql=sql & ",namxuatban='"&namxuatban&"'"
			sql=sql & ",het=" & hethang
			sql=sql & " WHERE NewsId=" & NewsId

			rs.open sql,con,1
			
			sql1="delete NewsDistribution where NewsId=" & NewsId & ";"
			if trim(request.Form("categoryid"))<>"" and sCatId=0 then
				for i=1 to UBound(ArrCat)
					sql1=sql1 & "insert into NewsDistribution "
					sql1=sql1 & "(NewsId,CategoryId) values "
					sql1=sql1 & "(" & NewsId
					sql1=sql1 & "," & ArrCat(i) & ")"
				Next
				CatId=ArrCat(1)
			else
				sql1=sql1 & "insert NewsDistribution "
				sql1=sql1 & "(NewsId,CategoryId) values "
				sql1=sql1 & "(" & NewsId
				sql1=sql1 & "," & CatId & ")"
			end if
			'response.write "sqlNewsDistribution=" & sql1
			rs.open sql1,con,1
			'set request.form=nothing
			set rs=nothing
			con.close
			set con=nothing
		   
		   if  smallPicture="Yes" then
            response.redirect ("news_nhapanh.asp?newsid=" & NewsId & "&catid=" & CatId & "&PictureCaption=" & PictureCaption & "&PictureAuthor=" & PictureAuthor & "&PictureAlign=" & PictureAlign & "&PictureDirection=" & PictureDirection & "&userr=" & session("user") & "&PictureId=" & old_PictureId)
			else
			response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
		    end if
		end if
	'set request.form=nothing
	set rs=nothing
	con.close
	set con=nothing
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
	Title_This_Page="Cập nhận tin & sản phẩm"
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
