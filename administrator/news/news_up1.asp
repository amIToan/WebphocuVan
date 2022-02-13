<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/include/Fs_liblary.asp" -->
<%
    LangID = Session("Language")
    if LangID = "" then LangID = "VN"
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
    	'Lay gia' tri CatId cua chuyen muc
			CatId=CLng(Upload.Form("CatId_DependRole"))
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
				sError=sError & "&nbsp;-&nbsp;Trạng thái gửi tin<br>"
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

        'IsSukien
		if Upload.Form("f_IsNew")<>"" and isnumeric(Upload.Form("f_IsNew")) then
			IsNew = CLng(Upload.Form("f_IsNew"))
		else
			IsNew = 0
		end if

		'AdsHome
		if Upload.Form("AdsHome")<>"" and isnumeric(Upload.Form("AdsHome")) then
			AdsHome = CLng(Upload.Form("AdsHome"))
		else
			AdsHome = 0
		end if

		'AdsNews
		''if Upload.Form("AdsNews")<>"" and isnumeric(Upload.Form("AdsNews")) then
		''	AdsNews = CLng(Upload.Form("AdsNews"))
		''else
		''	AdsNews = 0
		''end i
   
		'IsHotNews
		if Upload.form("IsHotNews")<>"" and isnumeric(Upload.form("IsHotNews")) then
			IsHotNews = CLng(Upload.form("IsHotNews"))
		else
			IsHotNews = 0
		end if

		PublicationNo=0
    
		'Title
	    Title=Trim(Upload.form("Title"))
        Desc = Upload.form("f_desc")
	    StoreOf=Trim(Upload.form("f_StoreOf"))

	    f_weight = Upload.form("f_Weight")
		
		Price=Upload.form("f_Price")
		if Price<> "" then
		    Price=chuan_money(Price)
		else
		    Price=0
		end if
        PriceNet=Upload.form("f_PriceNet")


		if PriceNet<> "" then
		    PriceNet=chuan_money(PriceNet)
		else
		    PriceNet=0
		end if
        F_Tskt =Trim(Upload.form("F_Tskt"))
        url_video =Trim(Upload.form("url_video"))
    
        'STdate =Trim(Upload.form("Datestart"))
        'ETdate =Trim(Upload.form("DateEnd"))


        'IF IsDate(STdate) THEN
        '    adate = Split(STdate,"-")
        '    Sdate = "'"&adate(1)&"/"&adate(0)&"/"&adate(2)&"'" 
        'ELSE
            Sdate  = "NULL"
        'END IF

        'IF IsDate(ETdate) THEN
        '   bdate = Split(ETdate,"-")
        '    Edate = "'"&bdate(1)&"/"&bdate(0)&"/"&bdate(2)&"'"     
        'ELSE
            Edate  = "NULL"
        'END IF
		'bodyx
		if (Trim(Upload.form("bodyx"))<>"") then
			body=Upload.form("bodyx")
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
		'Author
		Author=Trim(Upload.form("Author"))
		'Note
		IDCode=Upload.Form("f_IDCode")
		EmptyStore	=Upload.form("f_EmptyStore")
		f_Size	=Upload.form("f_Size")		
        Unit    =   Upload.Form("f_unit")
        FStatus    =   Upload.Form("StatusId")
        DateTem =  Upload.Form("DateCreater")
        meta_keyword =Trim(Upload.form("meta_keyword"))
        meta_desc    =Trim(Upload.form("meta_desc"))

        IF IsDate(DateTem) THEN          
           Idate = Split(DateTem,"-")          
             dd = Idate(0)
             mm = Idate(1)
             yy = Idate(2)
             DateCreater = mm&"/"&dd&"/"&yy
        ELSE
             DateCreater =  Now
        END  IF 

    if iStatus	=	"add" then
        PictureId=GetMaxId("Picture", "PictureId", "")
        NewsId=GetMaxId("News", "NewsId", "")
        sql="insert into News (NewsID, idcode, Title, Description,url_video,Author, DecsBannerImage, Body,IsHomeNews, IsSukien,IsNew, IsHotNews, PictureID , Price,PriceNet, Unit, StoreOf, Size,Weight, EmptyStore,meta_keyword,meta_desc, CreationDate, LastEditedDate, Status,languageid)"
        '                     (1     , 2     , 3    , 4          ,5        ,6     , 7              , 8   ,9         , 10      ,11   , 12       , 13        ,14    , 15     , 16  , 17     , 18  ,19    , 20        , 21         , 22      , 23          , 24            ,25     ,26 )"
        sql = sql & " values ("
        sql=sql & NewsId                            '1
        sql=sql & ",N'" & idcode & "'"              '2
        sql=sql & ",N'" & Title & "'"               '3
        sql=sql & ",N'" & Desc & "'"                '4
        sql=sql & ",N'" & url_video & "'"           '5
        sql=sql & ",N'" & Author & "'"              '6
        sql=sql & ",N'" & Trim(F_Tskt) & "'"        '7
        sql=sql & ",N'" & Trim(Body) & "'"          '8
        sql=sql & "," & IsHomeNews                  '9
        sql=sql & "," & IsSukien                    '10
        sql=sql & "," & IsHotNews                   '11
        sql=sql & "," & IsNew                       '12       
        sql=sql & "," & PictureID                   '13
       	sql=sql & ",'" & Price  &"'"				'14	
       	sql=sql & ",'" & PriceNet  &"'"				'15	
        sql=sql & ",N'" & unit  &"'"                '16
        sql=sql & ",N'" & StoreOf &"'"              '17
        sql=sql & ",N'" & f_Size & "'"              '18
        sql=sql & ",N'" & f_weight & "'"            '18
        sql=sql & ",'" & EmptyStore&"'"             '19
        sql=sql & ",N'" & meta_keyword & "'"        '20
        sql=sql & ",N'" & meta_desc & "'"           '21
        sql=sql & ",'" & DateCreater & "'"          '22
        sql=sql & ",'" & DateCreater & "'"          '23
        sql=sql & ",'" & FStatus& "'"  	            '24
        sql=sql & ",'" & LangID& "'"  		        '25
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

    'INSERT &  UPLOAD IMG
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
        Path=server.MapPath("/images_upload")
        'Anh dai dien      
        set FImgAv = Upload.Files("SmallPictureFileName")
        IF FImgAv Is Nothing THEN
        	FImgAv_=""
        ELSE
           Filetype = Right(FImgAv.FileName,len(FImgAv.Filename)-Instr(FImgAv.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FImgAv_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FImgAv_="P_"&dt&"1."&Filetype
        	   FImgAv.SaveAs Path &"\"&FImgAv_            
            END IF
        END IF
        'Anh slider
        set FImgLg = Upload.Files("LargePictureFileName")
        IF FImgLg Is Nothing THEN
        	FImgLg_=""
        ELSE
           Filetype = Right(FImgLg.FileName,len(FImgLg.Filename)-Instr(FImgLg.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FImgLg_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FImgLg_="P_"&dt&"2."&Filetype
        	   FImgLg.SaveAs Path &"\"&FImgLg_   
            END IF         
        END IF
        sqlP="insert into Picture (PictureId,SmallPictureFileName,LargePictureFileName,PictureAuthor,CategoryID,Creator,StatusID ) VALUES "
        sqlP=sqlP & "(" & PictureId
        sqlP=sqlP & ",'" &FImgAv_ & "'"
        sqlP=sqlP & ",'" &FImgLg_ & "'"
        sqlP=sqlP & ",N'adm'"
        sqlP=sqlP & ",'" &CatId&"'"		 							
        sqlP=sqlP & ",'adm','ed'"		 							
        sqlP=sqlP & ")"
        
        querry = sql&sql1&sqlP


        on error resume next
        con.Execute querry
        
        if err<>0 then
          response.write("Lỗi, xin vui lòng kiểm tra lại.")
        else
          response.write("Cập nhật thành công.")
        end if
              
    ELSE
        PicId= Upload.Form("old_PictureId")
        'sql="insert into News (NewsID, idcode, Title, Description,url_video,Author, DecsBannerImage, Body,IsHomeNews, IsSukien,IsNew, IsHotNews, PictureID , Price,PriceNet, Unit, StoreOf, Size,Width ,EmptyStore,meta_keyword,meta_desc, CreationDate, LastEditedDate, Status,languageid)"
        sql="UPDATE News SET "
        sql=sql & " idcode = N'"&idcode&"'"
        sql=sql & ",Title=N'"&Title&"'"
        sql=sql & ",Description=N'"&Desc& "'"
        sql=sql & ",url_video=N'"&url_video&"'"
        sql=sql & ",Author=N'" &Author& "'"
        sql=sql & ",DecsBannerImage =N'"&F_Tskt & "'"
        sql=sql & ",body=N'"& body&"'"
        sql=sql & ",IsHomeNews=" & IsHomeNews
        sql=sql & ",IsSukien="&IsSukien
        sql=sql & ",IsNew="&IsNew
        sql=sql & ",IsHotNews="&IsHotNews
        'sql=sql & ",PictureID='"&PictureID& "'"
        sql=sql & ",Price='"&Price&"'"
        sql=sql & ",PriceNet='"&PriceNet&"'"
        sql=sql & ",Unit='"&unit&"'"
        sql=sql & ",StoreOf=N'"& StoreOf & "'"
        sql=sql & ",Size=N'"& f_Size & "'"
        sql=sql & ",Weight=N'"& f_weight & "'"
        sql=sql & ",EmptyStore= N'" & EmptyStore&"'"
        sql=sql & ",meta_keyword=N'" & meta_keyword & "'"
        sql=sql & ",meta_desc=N'" & meta_desc & "'"
        sql=sql & ",LastEditedDate='" & getDateServer() & "'"  
        sql=sql & ",CreationDate='" & getDateServer() & "'" 
        sql=sql & ",StatusId='" & StatusId & "'"  
        sql=sql & ",Status='" & FStatus & "'"  
        sql=sql & ",LanguageID='" & LangID & "'"
        sql=sql & " WHERE NewsId=" & NewsId	
       
        set rs = Server.CreateObject("ADODB.Recordset")
        rs.open sql,con,1
        

                  
        Apicture  =   getColVal("Picture","SmallPictureFileName","PictureId = '"&PicId&"'")
        Spicture  =   getColVal("Picture","LargePictureFileName","PictureId = '"&PicId&"'")

        Path=server.MapPath("/images_upload")
        'Anh dai dien      
        set FImgAv = Upload.Files("SmallPictureFileName")
        set FImgSl = Upload.Files("LargePictureFileName")

        IF FImgAv Is Nothing And  FImgSl Is Nothing THEN
        	'2 File  Empty  > ko lam gi ca
        ELSE           
           sqlPic ="UPDATE Picture SET "  
           sqlPic=sqlPic & " Creator='AMC'" 
        '------------------------------------------------------------------------------------------------------------------------------------------------
           IF FImgAv Is Nothing THEN
                'FImgAv_="" 
           ELSE
           Filetype = Right(FImgAv.FileName,len(FImgAv.Filename)-Instr(FImgAv.Filename,"."))
                IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FImgAv_=""
                ELSE
                    dt = Trim(Replace(getDateServer(),"/",""))
                    dt = Trim(Replace(dt,":",""))
                    dt = Trim(Replace(dt," ",""))             
        	        FImgAv_="AV-"&dt&"."&Filetype
        	        FImgAv.SaveAs Path &"\"&FImgAv_    
    				sqlPic=sqlPic & ",SmallPictureFileName='" & FImgAv_ & "'"

                    'del file                
                    Uri1 = "/images_upload/"&Apicture&""                   
                    DelFile(Uri1)
                END IF 
            END IF'END  Avata
        '------------------------------------------------------------------------------------------------------------------------------------------------
           IF FImgSl Is Nothing THEN
                'FImgSl_="" 
           ELSE
           Filetype = Right(FImgSl.FileName,len(FImgSl.Filename)-Instr(FImgSl.Filename,"."))
                IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FImgSl_=""
                ELSE
                    dt = Trim(Replace(getDateServer(),"/",""))
                    dt = Trim(Replace(dt,":",""))
                    dt = Trim(Replace(dt," ",""))             
        	        FImgSl_="SL-"&dt&"."&Filetype
        	        FImgSl.SaveAs Path &"\"&FImgSl_     
					sqlPic=sqlPic & ",LargePictureFileName='" & FImgSl_ & "'"                  
                    'del file                
                    Uri2 = "/images_upload/"&Spicture&""                   
                    DelFile(Uri2)
                END IF 
            END IF'END  Avata
        '--------------------------------------------------------------------------------------------------------------------------------------------------           
			sqlPic=sqlPic & "  WHERE  PictureID='" & PicId & "'"                        
            set RsPic = Server.CreateObject("ADODB.Recordset")
            RsPic.open sqlPic,con,1    
        END IF '  END  FILE  Not Empty



       	


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

