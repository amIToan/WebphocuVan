<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
	'Kiem tra du lieu nhap vao
	Dim Upload 'Su dung AspUpload
	Dim sError
	sError="" 'Luu ca'c thong ba'o loi
	Set Upload = Server.CreateObject("Persits.Upload")

'	Upload.SetMaxSize 100000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save
	iStatus	=	Upload.Form("iStatus")	

    Response.Write iStatus

	if iStatus	=	"edit" then
		if Upload.Form("NewsId")<>"" and isnumeric(Upload.Form("NewsId")) then
			NewsId = CLng(Upload.Form("NewsId"))
		else
			NewsId = 0
		end if
		'sCatId
		if Upload.Form("sCatId")<>"" and isnumeric(Upload.Form("sCatId")) then
			sCatId = CLng(Upload.Form("sCatId"))
		else
			sCatId = 0
		end if
		
		'Kiem tra xem Gui tin len mot hay nhieu chuyen muc
		if trim(Upload.Form("categoryid"))="" then

			'Lay gia' tri CatId cua chuyen muc
			CatId=CLng(Upload.Form("CatId_DependRole"))
			'Kiem tra xem Chuyen muc na`y co' phu` hop voi StatusId khong
			if not IsNumeric(Upload.Form("StatusId")) then
				StatusId=0
			else
				StatusId=Clng((Upload.Form("StatusId")))
			end if
			
			if StatusId=0 then
				sError=sError & "Trạng thái gửi tin<br>"
			else
				StatusId=trim(CheckStatusWithCategoryId(NewsId,CatId,StatusId,Session("LstRole")))
				if len(StatusId)<>4 then
					sError=sError & "&nbsp;-&nbsp; " & StatusId & "<br>"
				end if
			end if
		elseif sCatId<>0 then

			if not IsNumeric(Upload.Form("StatusId")) then
				StatusId=0
			else
				StatusId=Clng((Upload.Form("StatusId")))
			end if
			CatId=sCatId
			if StatusId=0 then
				sError=sError & "&nbsp;-&nbsp; Tráº¡ng thÃ¡i gá»­i tin<br>"
			elseif StatusId=4 then
				sCatId=0
			else
				StatusId=trim(CheckStatusWithCategoryId(NewsId,CatId,StatusId,Session("LstRole")))
				if len(StatusId)<>4 then
					sError=sError & "&nbsp;-&nbsp; " & StatusId & "<br>"
				end if
			end if
		end if

		if trim(Upload.Form("categoryid"))<>"" and sCatId=0 then

			LstCategoryId=Trim(Upload.Form("categoryid"))
			ArrCat=Split(" " & LstCategoryId)
			CatId=ArrCat(1)
		end if
	else
		if trim(Upload.Form("categoryid"))="" then
		'Gui tin len mot chuyen muc
			'Lay gia' tri CatId cua chuyen muc
			CatId=CLng(Upload.Form("CatId_DependRole"))
			'Kiem tra xem Chuyen muc na`y co' phu` hop voi StatusId khong
			if not IsNumeric(Upload.Form("StatusId")) then
				StatusId=0
			else
				StatusId=Clng((Upload.Form("StatusId")))
			end if
		else
		'Gui tin len nhieu chuyen muc
			LstCategoryId=Trim(Upload.Form("categoryid"))
			ArrCat=Split(" " & LstCategoryId)
			CatId=ArrCat(1)
		end if	
	
	end if
		'Xa'c dinh xem User co' quyen gi` trong Chuyen muc na`y
		'Editor,GroupSenior,Approver,Administrator
		strTmp=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
		if trim(Upload.Form("categoryid"))<>"" then
		'Lay StatusId voi truong hop ban tin va`o nhieu chuyen muc
			statusId=strTmp & strTmp
		end if

		'IsHomeNews
		if Upload.Form("IsHomeNews")<>"" and isnumeric(Upload.Form("IsHomeNews")) then
			IsHomeNews = CLng(Upload.Form("IsHomeNews"))
		else
			IsHomeNews = 0
		end if

        'IsSukien
		if Upload.Form("IsSukien")<>"" and isnumeric(Upload.Form("IsSukien")) then
			IsSukien = CLng(Upload.Form("IsSukien"))
		else
			IsSukien = 0
		end if

        'IsShare
		if Upload.Form("IsShare")<>"" and isnumeric(Upload.Form("IsShare")) then
			IsShare = CLng(Upload.Form("IsShare"))
		else
			IsShare = 0
		end if

		'AdsHome
		if Upload.Form("AdsHome")<>"" and isnumeric(Upload.Form("AdsHome")) then
			AdsHome = CLng(Upload.Form("AdsHome"))
		else
			AdsHome = 0
		end if

		'AdsNews
		if Upload.Form("AdsNews")<>"" and isnumeric(Upload.Form("AdsNews")) then
			AdsNews = CLng(Upload.Form("AdsNews"))
		else
			AdsNews = 0
		end if

		'EventId
		if isNumeric(Upload.form("eventid")) then
			EventId=CLng(Upload.form("eventid"))
		else
			EventId=0
		end if
		'IsCatHomeNews
		if Upload.form("IsCatHomeNews")<>"" and isnumeric(Upload.form("IsCatHomeNews")) then
			IsCatHomeNews = CLng(Upload.form("IsCatHomeNews"))
		else
			IsCatHomeNews = 0
		end if
		'IsCatHomeNews_Below
		
        IsCatHomeNews_Below = Replace(Upload.form("txtKho"),"'","''")
		'languageid
		if len(Upload.form("languageid"))>1 then
			languageid=Trim(Upload.form("languageid"))
		else
			languageid="VN"
		end if
		'IsHotNews
		if Upload.form("IsHotNews")<>"" and isnumeric(Upload.form("IsHotNews")) then
			IsHotNews = CLng(Upload.form("IsHotNews"))
		else
			IsHotNews = 0
		end if

		'IsSlaveHomePageNews
		if Upload.form("sltVAT")<>"" and isnumeric(Upload.form("sltVAT")) then
			IsSlaveHomePageNews = CLng(Upload.form("sltVAT"))
		else
			IsSlaveHomePageNews = 0
		end if
		
		PublicationNo=0
		'Title
		if Trim(Upload.form("Title"))<>"" then
			Title=Trim(replace(Upload.form("Title"),"'","''"))
			Title=Replace(Title,"""","&quot;")
		else
			sError=sError & "&nbsp;-&nbsp; TiÃªu Ä‘á»<br>"
		end if
		if Trim(Upload.form("f_StoreOf"))<>"" then
			StoreOf=Trim(Upload.form("f_StoreOf"))
			StoreOf=Replace(StoreOf,"'","''")
			StoreOf=Replace(StoreOf,chr(13) & chr(10),"<br>")
		else
			StoreOf=""
		end if
		
		if Upload.form("f_Weight")<>"" and isnumeric(Upload.form("f_Weight")) then
			Weight = CLng(Upload.form("f_Weight"))
		else
			Weight = 0
		end if			

		Price=Upload.form("f_Price")
		if Price<> "" then
			Price=chuan_money(Price)
		else
			Price=0
		end if

		'Description
		
        if Trim(Upload.form("Description"))<>"" then
			Description=Trim(replace(Upload.form("Description"),"'","''"))
			Description=Replace(Description,"""","&quot;")	
		end if

        txtTT =Trim(Upload.form("txtTT"))
        F_Tskt =Trim(Upload.form("F_Tskt"))
        cauhinh =Trim(Upload.form("Fconfig"))

		'bodyx
		if (Trim(Upload.form("bodyx"))<>"") then
			body= Upload.form("bodyx")
		else
			sError=sError & "&nbsp;-&nbsp; Nội dung<br>"
		end if

      
		'PictureAlign
		PictureAlign=Trim(Upload.form("PictureAlign"))
		'PictureDirection
		PictureDirection=Trim(Upload.form("PictureDirection"))
		if not IsNumeric(PictureDirection) or PictureDirection="" then
			PictureDirection=0
		else
			PictureDirection=1
		end if
		'PictureCaption
		PictureCaption=Trim(Upload.form("PictureCaption"))
		PictureCaption=Replace(PictureCaption,"'","''")
		PictureCaption=Replace(PictureCaption,"""","&quot;")
		'PictureAuthor
		PictureAuthor=Trim(Upload.form("PictureAuthor"))
		PictureAuthor=Replace(PictureAuthor,"'","''")
		PictureAuthor=Replace(PictureAuthor,"""","&quot;")
		'Author
		Author=Trim(Upload.form("Author"))
		Author=Replace(Author,"'","''")
		Author=Replace(Author,"""","&quot;")
		'Source
		source=Upload.Form("txtBH")
		'Note

		IDCode=Upload.Form("f_IDCode")

		EmptyStore	=Upload.form("f_EmptyStore")


		f_Size	=Upload.form("f_Size")

		
        Unit    =   Upload.Form("f_unit")
        Unit    =   Upload.Form("f_unit")
        FStatus    =   Upload.Form("StatusId")

        DateTem =  Upload.Form("DateCreater")
        DateTem1 =  Upload.Form("SDate")
        DateTem2 =  Upload.Form("EDate")

        meta_keyword =Trim(Upload.form("meta_keyword"))
        meta_desc    =Trim(Upload.form("meta_desc"))

        url_video    =Trim(Upload.form("url_video"))

        IF IsDate(DateTem) THEN 
            
           Idate = Split(DateTem,"-")
            
             dd = Idate(0)
             mm = Idate(1)
             yy = Idate(2)
             DateCreater = mm&"/"&dd&"/"&yy
        ELSE
             DateCreater =  Now
        END  IF 


  '      IF IsDate(DateTem1) THEN 
  '          
  '         Idate1 = Split(DateTem1,"-")
  '          
  '           dd = Idate1(0)
  '           mm = Idate1(1)
  '           yy = Idate1(2)
  '           SDate = mm&"/"&dd&"/"&yy
  '      END  IF 
            SDate="NULL"
            
 '       IF IsDate(DateTem2) THEN 
 '           
 '          Idate2 = Split(DateTem2,"-")
 '           
 '            dd = Idate2(0)
 '            mm = Idate2(1)
 '            yy = Idate2(2)
 '            EDate = mm&"/"&dd&"/"&yy
 '       END  IF 
            EDate="NULL"
		if iStatus = "add"	 then
			Note=Trim(Upload.form("Note"))
			'Note=Replace(Note,"'","''")
			'Note=Replace(Note,chr(13) & chr(10),"<br>")
			Note=old_Note & " Post by: <b>" & session("user") & "</b> At: " & Hour(now) & "h" & Minute(now) & """&nbsp;" & Day(now) & "/" & Month(now) & "/" & Year(now) & "<br>" & Note
		else
    
			if Upload.Form("old_PictureId")<>"" then
				old_PictureId = CLng(Upload.Form("old_PictureId"))
			else
				old_PictureId = 0
			end if



			'old_Note
			old_Note=Trim(Upload.form("old_Note"))
			old_Note=Replace(old_Note,"'","''")
			old_Note=Replace(old_Note,chr(13) & chr(10),"<br>")
			'Note
			Note=Trim(Upload.form("Note"))
			'Note=Replace(Note,"'","''")
			'Note=Replace(Note,chr(13) & chr(10),"<br>")
			Note=old_Note & " Post by: <b>" & session("user") & "</b> At: " & Hour(now) & "h" & Minute(now) & """&nbsp;" & Day(now) & "/" & Month(now) & "/" & Year(now)
			

	''''	'RemoveImage
	''''	if Upload.Form("RemoveImage")<>"" and isnumeric(Upload.Form("RemoveImage")) then
	''''		RemoveImage = CLng(Upload.Form("RemoveImage"))
	''''	else
	''''		RemoveImage = 0
	''''	end if
	''''	'RemoveLargeImage
	''''	if Upload.Form("RemoveLargeImage")<>"" and isnumeric(Upload.Form("RemoveLargeImage")) then
	''''		RemoveLargeImage = CLng(Upload.Form("RemoveLargeImage"))
	''''	else
	''''		RemoveLargeImage = 0
	''''	end if
			
			
		End if
		'smallpicturefilename
        PicTitle    =   Uni2NONE(Title)
        PicTitle    =   replace(PicTitle," ","_")
        PicTitle    =   Replace(PicTitle,"-","_")

        if Len(PicTitle)>=100 then
            PicTitle    =   left(left,99)
        end if


		set smallpicture = Upload.Files("SmallPictureFileName")
		If smallpicture Is Nothing Then
			SmallPictureFileName=""
		else
		   Filetype = Right(smallpicture.Filename,len(smallpicture.Filename)-Instr(smallpicture.Filename,"."))
            if old_PictureId=0 then
				PictureId=GetMaxId("Picture", "PictureId", "")
				SmallPictureFileName="small_" & PictureId & "." & Filetype
			else
				SmallPictureFileName="small_" & old_PictureId & "." & Filetype
			end if

		'	if old_PictureId=0 then
		'		PictureId=GetMaxId("Picture", "PictureId", "")
		'		SmallPictureFileName=PicTitle & PictureId & "." & Filetype
		'	else
		'		SmallPictureFileName=PicTitle & old_PictureId & "." & Filetype
		'	end if
		End If

		'largepicturefilename
		set largepicture = Upload.Files("LargePictureFileName")
		If largepicture Is Nothing Then
			LargePictureFileName=""
		else
		   Filetype = Right(largepicture.Filename,len(largepicture.Filename)-Instr(largepicture.Filename,"."))
			if old_PictureId=0 then
				LargePictureFileName="large_" & PictureId & "." & Filetype
			else
				LargePictureFileName="large_" & old_PictureId & "." & Filetype
			end if
		End If	
		'Response.Write("SmallPictureFileName"&Path & "\" & SmallPictureFileName)



		Dim ar_picturefile(16)	
		Dim ar_picturename(16) 
		dim ar_contentPicture(16)
		for n = 1 to 16 
			set ar_picturefile(n) = Upload.Files("PictureFile"&n)
			ar_contentPicture(n)	=	Trim(Upload.form("ContentPicture"&n))
			'Response.Write("<br>ar_contentPicture("&n&")"&ar_contentPicture(n))
			If ar_picturefile(n) Is Nothing Then
				ar_picturename(n)=""
			else	
			   	Filetype = Right(ar_picturefile(n).Filename,len(ar_picturefile(n).Filename)-Instr(ar_picturefile(n).Filename,"."))
				if Lcase(Filetype)="jpg" or Lcase(Filetype)="gif" or Lcase(Filetype)="jpeg" or Lcase(Filetype)="png"then
					filename	=	"Picture_"
				else
					filename	=	"Doccument_"
				end if
				if old_PictureId=0 then
					ar_picturename(n)=filename& n &"_" & PictureId & "." & Filetype
				else
					ar_picturename(n)=filename& n &"_" & old_PictureId & "." & Filetype
				end if		
			end if			
		next

			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")

     if iStatus	=	"add" then
			 	if CLng(PictureId)>0 then
			 	'BÃ¢y giá» cÃ³ áº£nh
			 		smallpicture.SaveAs Path & "\" & SmallPictureFileName
			 		if LargePictureFileName<>"" then
			 			largepicture.SaveAs Path & "\" & LargePictureFileName
			 		end if
			 		for n=1 to 16 
			 			if ar_picturename(n) <>"" then
			 				Filetype = Right(ar_picturefile(n).Filename,len(ar_picturefile(n).Filename)-Instr(ar_picturefile(n).Filename,"."))
			 				if Lcase(Filetype)="jpg" or Lcase(Filetype)="gif" or Lcase(Filetype)="jpeg" or Lcase(Filetype)="png" then
			 					ar_picturefile(n).SaveAs Path & "\" & ar_picturename(n)
			 				else
			 					ar_picturefile(n).SaveAs PathDoc & "\" & ar_picturename(n)
			 				end if						
			 			end if
			 		next
			 		sql="insert into Picture (PictureId,PictureCaption,SmallPictureFileName,"
			 		sql=sql & "LargePictureFileName,PictureAuthor,Creator,CategoryID,StatusID"
                     for n=1 to 16
                         if ar_picturename(n) <> "" then
                             sql=sql & ",PictureFile"&n&",ContentPicture"&n
                         end if
                     next
             
                     sql=sql & ") values "
			 		sql=sql & "(" & PictureId
			 		sql=sql & ",N'" & PictureCaption & "'"
			 		sql=sql & ",'" & SmallPictureFileName & "'"
			 		sql=sql & ",'" & LargePictureFileName & "'"
			 		sql=sql & ",N'" & PictureAuthor & "'"
			 		sql=sql & ",N'" & session("user") & "'"
			 		sql=sql & ",'" & CatId&"','ed'"
			 		for n=1 to 16
                         if ar_picturename(n) <> "" then
			 			    sql=sql &",'"& ar_picturename(n) &"'"
			 			    sql=sql &",N'"&ar_contentPicture(n) &"'"
                         end if
			 		next					
			 		sql=sql & ")"
			 		response.write "sqlPicture=" & sql & "<br>"
			 		rs.open sql,con,1
			 	else
			 	'BÃ¢y giá» cÅ©ng khÃ´ng cÃ³ áº£nh
			 		PictureId=0
			 	end if
	end if    
    		
    if iStatus	=	"edit" then
			
			set rs=server.CreateObject("ADODB.Recordset")
             PictureId =  Upload.Form("old_PictureId")
    if PictureId <> 0 then
					sql="Update Picture set "
					sql=sql & "PictureCaption=N'" & PictureCaption & "'"
					if SmallPictureFileName<>"" then
						smallpicture.SaveAs Path & "\" &  SmallPictureFileName
						sql=sql & ",SmallPictureFileName='" & SmallPictureFileName & "'"
					end if
					if RemoveLargeImage=1 then
						sql=sql & ",LargePictureFileName=''"
					elseif LargePictureFileName<>"" then
						largepicture.SaveAs Path & "\" &  LargePictureFileName
						sql=sql & ",LargePictureFileName='" & LargePictureFileName & "'"
					end if
					sql=sql & ",PictureAuthor=N'" & PictureAuthor & "'"
					sql=sql & ",CategoryID=" & CatId
					for n=1 to 16
						if ar_picturename(n) <>"" then
							Filetype = Right(ar_picturefile(n).Filename,len(ar_picturefile(n).Filename)-Instr(ar_picturefile(n).Filename,"."))
							if Lcase(Filetype)="jpg" or Lcase(Filetype)="gif" or Lcase(Filetype)="jpeg" or Lcase(Filetype)="png"then
								ar_picturefile(n).SaveAs Path & "\" & ar_picturename(n)
							else
								ar_picturefile(n).SaveAs PathDoc & "\" & ar_picturename(n)
							end if										
							sql=sql &",PictureFile"&n&" ='"& ar_picturename(n) &"'"
							sql=sql &",ContentPicture"&n&" =N'"& ar_contentPicture(n) &"'"
						elseif GetNumeric(Upload.form("PictureDel"&n),0) = 1 then
							sql=sql &",PictureFile"&n&" ='' "
							sql=sql &",ContentPicture"&n&" ='' "
						else
							sql=sql &",ContentPicture"&n&" =N'"& ar_contentPicture(n) &"'"
						end if				
					next					
					sql=sql & " where PictureId=" & PictureId
					Response.Write("<br>"&sql&"<br>")
					rs.open sql,con,1
            
		else
     PicTitle    =   Uni2NONE(Title)
        PicTitle    =   replace(PicTitle," ","_")
        PicTitle    =   Replace(PicTitle,"-","_")
        if Len(PicTitle)>=100 then
            PicTitle    =   left(left,99)
        end if


		set smallpicture = Upload.Files("SmallPictureFileName")
		If smallpicture Is Nothing Then
			SmallPictureFileName=""
		else
		   Filetype = Right(smallpicture.Filename,len(smallpicture.Filename)-Instr(smallpicture.Filename,"."))
			if old_PictureId=0 then
				PictureId=GetMaxId("Picture", "PictureId", "")
				SmallPictureFileName=PicTitle & PictureId & "." & Filetype
			else
				SmallPictureFileName=PicTitle & old_PictureId & "." & Filetype
			end if
		End If

		'largepicturefilename
		set largepicture = Upload.Files("LargePictureFileName")
		If largepicture Is Nothing Then
			LargePictureFileName=""
		else
		   Filetype = Right(largepicture.Filename,len(largepicture.Filename)-Instr(largepicture.Filename,"."))
			if old_PictureId=0 then
				LargePictureFileName="large_" & PictureId & "." & Filetype
			else
				LargePictureFileName="large_" & old_PictureId & "." & Filetype
			end if
		End If	
		'Response.Write("SmallPictureFileName"&Path & "\" & SmallPictureFileName)
		for n = 1 to 16 
			set ar_picturefile(n) = Upload.Files("PictureFile"&n)
			ar_contentPicture(n)	=	Trim(Upload.form("ContentPicture"&n))
			'Response.Write("<br>ar_contentPicture("&n&")"&ar_contentPicture(n))
			If ar_picturefile(n) Is Nothing Then
				ar_picturename(n)=""
			else	
			   	Filetype = Right(ar_picturefile(n).Filename,len(ar_picturefile(n).Filename)-Instr(ar_picturefile(n).Filename,"."))
				if Lcase(Filetype)="jpg" or Lcase(Filetype)="gif" or Lcase(Filetype)="jpeg" or Lcase(Filetype)="png"then
					filename	=	"Picture_"
				else
					filename	=	"Doccument_"
				end if
				if old_PictureId=0 then
					ar_picturename(n)=filename& n &"_" & PictureId & "." & Filetype
				else
					ar_picturename(n)=filename& n &"_" & old_PictureId & "." & Filetype
				end if		
			end if			
		next
    if CLng(PictureId)>0 then
			 	'bay gio co anh
			 		smallpicture.SaveAs Path & "\" & SmallPictureFileName
			 		if LargePictureFileName<>"" then
			 			largepicture.SaveAs Path & "\" & LargePictureFileName
			 		end if
			 		for n=1 to 16 
			 			if ar_picturename(n) <>"" then
			 				Filetype = Right(ar_picturefile(n).Filename,len(ar_picturefile(n).Filename)-Instr(ar_picturefile(n).Filename,"."))
			 				if Lcase(Filetype)="jpg" or Lcase(Filetype)="gif" or Lcase(Filetype)="jpeg" or Lcase(Filetype)="png" then
			 					ar_picturefile(n).SaveAs Path & "\" & ar_picturename(n)
			 				else
			 					ar_picturefile(n).SaveAs PathDoc & "\" & ar_picturename(n)
			 				end if						
			 			end if
			 		next
			 		sql="insert into Picture (PictureId,PictureCaption,SmallPictureFileName,"
			 		sql=sql & "LargePictureFileName,PictureAuthor,Creator,CategoryID,StatusID"
                     for n=1 to 16
                         if ar_picturename(n) <> "" then
                             sql=sql & ",PictureFile"&n&",ContentPicture"&n
                         end if
                     next
             
                     sql=sql & ") values "
			 		sql=sql & "(" & PictureId
			 		sql=sql & ",N'" & PictureCaption & "'"
			 		sql=sql & ",'" & SmallPictureFileName & "'"
			 		sql=sql & ",'" & LargePictureFileName & "'"
			 		sql=sql & ",N'" & PictureAuthor & "'"
			 		sql=sql & ",N'" & session("user") & "'"
			 		sql=sql & ",'" & CatId&"','ed'"
			 		for n=1 to 16
                         if ar_picturename(n) <> "" then
			 			    sql=sql &",'"& ar_picturename(n) &"'"
			 			    sql=sql &",N'"&ar_contentPicture(n) &"'"
                         end if
			 		next					
			 		sql=sql & ")"
			 		response.write "sqlPicture=" & sql & "<br>"
			 		rs.open sql,con,1
			 	else
			 	'BÃ¢y giá» cÅ©ng khÃ´ng cÃ³ áº£nh
			 		PictureId=0
			 	end if
    end if
				
				'Response.End()		

   End if
    if iStatus	=	"add" then
				NewsId=GetMaxId("News", "NewsId", "")
				sql="insert into News (NewsID, idcode,url_video, Title, Description,Author, DecsBannerImage, Body,IsHomeNews, IsCatHomeNews,IsSukien,IsShare"
                sql = sql & " ,IsCatHomeNews_Below, IsHotNews, AdsHome, AdsNews, PictureID, Source , Price, Unit, StoreOf, Size, EmptyStore, Note,meta_keyword,meta_desc, Creator, CreationDate, LastEditor, LastEditedDate, Status)"
				sql = sql & " values ("
				sql=sql & NewsId
				sql=sql & ",N'" & idcode & "'"

                sql=sql & ",N'" & url_video & "'"

				sql=sql & ",N'" & Title & "'"
				sql=sql & ",N'" & Description & "'" 'Mo ta
				sql=sql & ",N'" & Trim(cauhinh) & "'"  ' mo ta cau hinh (author)
			
                sql=sql & ",N'" & Trim(F_Tskt) & "'" 'thong so ky thuat
				sql=sql & ",N'" & Trim(Body) & "'"
				sql=sql & "," & IsHomeNews
				sql=sql & "," & IsCatHomeNews 'Loai hang  (cu, moi)

                sql=sql & "," & IsSukien 'IsSukien
                sql=sql & "," & IsShare 'IsShare

				sql=sql & ",N'" & IsCatHomeNews_Below&"'" ' Kho (con , het)
				sql=sql & "," & IsHotNews
				sql=sql & "," & AdsHome
				sql=sql & "," & AdsNews								
				sql=sql & "," & PictureID
				sql=sql & ",N'" & source & "'"  ' Bao hanh
                sql=sql & ",'" & Price  &"'"
                sql=sql & ",N'" & unit  &"'"
                sql=sql & ",N'" & StoreOf &"'"
                sql=sql & ",N'" & f_Size & "'"
                sql=sql & ",'" & EmptyStore&"'"
                sql=sql & ",N'" & Note & "'"
                sql=sql & ",N'" & meta_keyword & "'"
                sql=sql & ",N'" & meta_desc & "'"
                sql=sql & ",N'" & session("user") & "'" 'Creator
                sql=sql & ",'" & DateCreater & "'" ' CreationDate   
				sql=sql & ",N'" & session("user") & "'" 'LastEditor
                sql=sql & ",'" & DateCreater & "'" 'LastEditedDate
        '        sql=sql & ",'" & SDate & "'" 'SDate
        '       sql=sql & ",'" & EDate & "'" 'EDate
				sql=sql & ",'" & FStatus& "'" 'StatusId	
				sql = sql & ")"
		if trim(Upload.form("categoryid"))<>"" then
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
			'rs.open sql1,con,1
			''set Upload.form=nothing
			'set rs=nothing
			'con.close
			'set con=nothing
    querry = sql&sql1
    on error resume next
            con.Execute querry

  if err<>0 then
    response.write("Lỗi, xin vui lòng kiểm tra lại.")
  else
    response.write("Cập nhật thành công.")
  end if

		else
			sql="Update News set "
			sql=sql & " idcode = N'" & idcode & "'"
			sql=sql & ",Title=N'" & Title & "'"
			sql=sql & ",Description=N'" & Description & "'"
            sql=sql & ",DecsBannerImage =N'" & F_Tskt & "'"
			sql=sql & ",body=N'" & body & "'"
			sql=sql & ",Author=N'" & cauhinh & "'"
			sql=sql & ",Source=N'" & Source & "'"
			sql=sql & ",IsHomeNews=" & IsHomeNews
			sql=sql & ",IsCatHomeNews=" & IsCatHomeNews

            sql=sql & ",IsSukien=" & IsSukien 'IsSukien
            sql=sql & ",IsShare=" & IsShare 'IsShare
        
			sql=sql & ",IsCatHomeNews_Below=N'" & IsCatHomeNews_Below&"'"
			sql=sql & ",IsHotNews=" & IsHotNews
			sql=sql & ",IsSlaveHomePageNews=" & IsSlaveHomePageNews
			sql=sql & ",AdsHome='"&AdsHome&"'"	
			sql=sql & ",AdsNews='"&AdsNews&"'"
			sql=sql & ",Price='"&Price&"'"			
			sql=sql & ",StoreOf=N'"& StoreOf & "'"
			sql=sql & ",PictureID=" & PictureID
			sql=sql & ",PictureAlign='" & PictureAlign & "'"
			sql=sql & ",LanguageID='" & LanguageID & "'"
			sql=sql & ",PublicationNo=" & PublicationNo 
			sql=sql & ",LastEditor=N'" & userr & "'" 'LastEditor
        
            sql=sql & ",url_video=N'" & url_video & "'" 'LastEditor

		'	sql=sql & ",LastEditedDate='" & getDateServer() & "'" 'LastEditedDate
            sql=sql & ",LastEditedDate='" & DateCreater & "'" 'LastEditedDate
			sql=sql & ",CreationDate='" & DateCreater & "'" 'LastEditedDate
       '     sql=sql & ",SDate='" & SDate & "'" 'SDate
       '     sql=sql & ",EDate='" & EDate & "'" 'EDate
			sql=sql & ",StatusId='" & StatusId & "'" 'StatusId
			sql=sql & ",Status='" & FStatus & "'" 'StatusId
			if Editor<>"" then
				sql=sql & ",Editor='" & Editor & "'"
			end if
			if GroupSenior<>"" then
				sql=sql & ",GroupSenior='" & GroupSenior & "'"
			end if
			if Approver<>"" then
				sql=sql & ",Approver='" & Approver & "'"
			end if
            sql=sql & ",meta_keyword=N'" & meta_keyword & "'"
			sql=sql & ",meta_desc=N'" & meta_desc & "'"

			sql=sql & ",Note=N'" & Note & "'"
			sql=sql & ",Size='"&f_Size&"'"
			sql=sql & ",Weight='"&Weight&"'"
			sql=sql & ",EmptyStore= N'" & EmptyStore&"'"
			sql=sql & " WHERE NewsId=" & NewsId		
			Response.Write("sql : " &sql)
			rs.open sql,con,1
			
			sql1="delete NewsDistribution where NewsId=" & NewsId & ";"
			if trim(Upload.Form("categoryid"))<>"" and sCatId=0 then
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
			rs.open sql1,con,1	
		end if
		if GetNumeric(Upload.form("attach_product"),0) = 1 then
			response.redirect ("attach_news.asp?newsid=" & NewsId & "&catid=" & CatId)
		else
			response.redirect ("news_insertsuccess.asp?newsid=" & NewsId & "&catid=" & CatId)
		end if
			Response.End()
	set Upload=nothing
	set rs=nothing
	con.close
	set con=nothing
%>

