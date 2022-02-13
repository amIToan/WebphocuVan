<!--#include virtual="/include/config.asp" -->
<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Function navigate(str_,mid_,id_) 
        link =""
        IF str_ <> "" THEN  link =  link&"/"&str_       
        IF mid_ <>  "" THEN  link =  link&"/"&mid_
        IF id_ <> "" THEN  link =  link&"/"&id_
        IF link <>  "" THEN  navigate =  link ELSE navigate = "#"
    End Function 
%>



<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getLink(cateid,newsid,title) 
    link = ""
    if cateid <> "" then link = link & "/"&cateid
    if newsid <> "" then link = link & "/"&newsid
    if title <> "" then link = link & "/"& Replace(Uni2NONE(LCase(title))," ","-")&".html"
    if link <> "" then getLink = link else getLink = "#"
end function 
%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function LinkUri(cateid,newsid,title) 
    link = ""
    if cateid <> "" then link = link & "/"&cateid
    if newsid <> "" then link = link & "/"&newsid
    if title <> "" then link = "/"& Replace(Uni2NONE(title)," ","-")&link&".html"
    if link <> "" then LinkUri = link else LinkUri = "#"
end function 
%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getColVal(tabName,colName,query) 
    if tabName = "" or colName = "" or query = "" then
        getColVal = ""
    else
        sqlCV = "Select top 1 "&colName&" from "&tabName&" where "&query
        set rsCV = server.CreateObject("ADODB.RecordSet")
        'Response.Write sqlCV
        rsCV.open sqlCV,con,1
            if not rsCV.EOF then
                getColVal = Trim(rsCV(colName))
            else
                getColVal = ""
            end if
        set rsCV = nothing
    end if
End Function 
%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function IsNews(NID) 
    IF NID = "" THEN
        IsNews = -1
    ELSE
       '(SDate IS NULL) AND  (EDate   IS  NULL)
        sqlx = "SELECT  TOP 1 * FROM V_News WHERE NewsID = '"&NID&"' AND CategoryLoai<>4" 
        set Rsx = server.CreateObject("ADODB.RecordSet")
        response.write sqlx
        Rsx.open sqlx,con,1
            if not Rsx.EOF then
                IsNews =  1
            else
                IsNews =  0
            end if
        set Rsx = nothing
    END IF
End Function 
%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Function Record_Count(tb,dk) 
        Total = 0
    IF tb = "" THEN
        Total = 0
    ELSE
        dkien   = " WHERE "&dk
        sqlx = "SELECT COUNT(NewsID) AS TotalRecord FROM "&tb&"  "&dkien
        set Rsx = server.CreateObject("ADODB.RecordSet")
        Rsx.open sqlx,con,1
            if not Rsx.EOF then
                Total =  Rsx("TotalRecord")
            else
                Total =  0
            end if
        set Rsx = nothing
    END IF
    Record_Count = Total
End Function 
%>





<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Function Is_Mobile()
        Set Regex = New RegExp
        With Regex
          .Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad)"
          .IgnoreCase = True
          .Global = True
        End With
        Match = Regex.test(Request.ServerVariables("HTTP_USER_AGENT"))
        If Match then
          Is_Mobile = True
        Else
          Is_Mobile = False
        End If
    End Function
%>


<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Function GetListParentCat(cid_)



            ''''IF Not rs1.EOF THEN
            ''''
            ''''Temp_  = rs1("ParentCategoryId")
            ''''
            ''''IF Not IsEmpty(Temp_) And  IsNumeric(Temp_) THEN
            ''''
            ''''    PCatId=Cint(rs1("ParentCategoryId"))
            ''''    if PcatId<>0 then
            ''''    	i=i+1
            ''''    	ArrValue(i)=rs1("ParentCategoryId")
            ''''    end if           
            ''''ELSE
            ''''    
            ''''END  IF
            ''''
            ''''END  IF

    Response.Write "xxxxxxxxxxxxxxxxx:"&cid_

    IF  IsNumeric(CatId) THEN
	'Get Tree List CategoryId of Inpute Category.
	'Result is a string of CategoryId separated by spacebar, not include Input Category
	Dim i,ArrValue(100)
	i=0
	Dim rs1
	set rs1=Server.CreateObject("ADODB.Recordset")

	PCatId=CatId
    
	  Do while PCatId<>0
		sql_GetListParentCat="select ParentCategoryId from NewsCategory where CategoryId='" & PCatId&"'"
		rs1.open sql_GetListParentCat,con,1

			PCatId=Cint(rs1("ParentCategoryId"))
			if PcatId<>0 then
				i=i+1
				ArrValue(i)=rs1("ParentCategoryId")
			end if


		rs1.close
	  Loop
	GetListParentCat=Trim(Join(ArrValue))
    ELSE
        '----
    END IF
End Function%>



<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function GetFullDate(theDate,lang)
	ngay = Clng(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = Clng(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = Clng(year(theDate))
    
if lang="VN" then
	GetFullDate=ngay & "/" & thang & "/" & nam
else
	GetFullDate=MonthName(Month(now()),true) & " " & ngay & "," & nam
end if
End Function%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function GetFullDateTime(theDate)
	gio=Clng(Hour(theDate))
	if gio<10 then
		gio="0" & gio
	end if
	phut=Clng(Minute(theDate))
	if phut<10 then
		phut="0" & phut
	end if
	ngay = Clng(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = Clng(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = Clng(year(theDate))
	GetFullDateTime=gio & "h" & phut & """&nbsp;" & ngay & "/" & thang & "/" & nam
End Function%>



<%'VIETSOFT 03/08/2015 %>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function FomatDateTime(str_dateTime) 

    if str_dateTime <> "" and IsDate(str_dateTime) then 
     
    Ngay = Day(str_dateTime)

    if len(Ngay) = 1  then
        Ngay = "0"&Ngay 
    end if

    Thang= Month(str_dateTime)
    if len(Thang) = 1  then
        Thang = "0"&Thang 
    end if

    Nam =  Year(str_dateTime)
        
    Gio  = Hour(str_dateTime) 
    if len(Gio) = 1  then
        Gio = "0"&Gio 
    end if
    Phut = Minute(str_dateTime)
    if len(Phut) = 1  then
        Phut = "0"&Phut 
    end if
    Giay = Second(str_dateTime)   
        if len(str_dateTime) = 1  then
            Giay = "0"&Giay
        elseif len(Giay) = 1  then
             Giay = "00"
        end if
        
    if Hour(str_dateTime)  > 12 then
        SttNgay = "PM"
    elseif  Hour(str_dateTime)  <= 12 then
        SttNgay = "AM"
    end if

    FomatDateTime = Ngay &"-"& Thang &"-"& Nam &" "& Gio &"  :"& Phut &"  :"& Giay  
    else
        FomatDateTime ="-"
    end if
    end  Function
%>


<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Function FomatDateTime_sql(str_FullTime) 

    M= split(str_FullTime,"/")
    if str_FullTime <> "" and IsDate(str_FullTime) then 
     
    Ngay = M(0)

    if len(Ngay) = 1  then
        Ngay = "0"&Ngay 
    end if

    Thang=  M(1)
    if len(Thang) = 1  then
        Thang = "0"&Thang 
    end if

    Nam =  Year(str_FullTime)
        
    Gio  = Hour(str_FullTime) 
    if len(Gio) = 1  then
        Gio = "0"&Gio 
    end if
    Phut = Minute(str_FullTime)
    if len(Phut) = 1  then
        Phut = "0"&Phut 
    end if
    Giay = Second(str_FullTime)   
        if len(str_FullTime) = 1  then
            Giay = "0"&Giay
        elseif len(Giay) = 1  then
             Giay = "00"
        end if
        


    FomatDateTime_sql = Nam &"-"& Thang &"-"&Ngay  &" "& Gio &":"& Phut &":"& Giay  
    else
        FomatDateTime_sql ="-"
    end if
    end  Function
%>



<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  '+ - time
    Function DateTime_Total(TTime,donvi,BeginTime,EndTime)     
        IF BeginTime <> "" and  EndTime <> ""  and IsDate (BeginTime) and IsDate(EndTime) THEN  
             l_BeginTime = Replace(BeginTime, "/", "-")
             l_EndTime   = Replace(EndTime, "/", "-")
             Total_time = datediff("n",l_BeginTime,l_EndTime) ' in phút
             IF Total_time  < 0 THEN 
                b_Ngay  = Day(BeginTime)
                b_Thang = Month(BeginTime)
                b_Nam   = Year(BeginTime)
                b_Gio   = Hour(BeginTime)
                b_Phut  = Minute(BeginTime)
                b_Giay  = Second(BeginTime)

                e_Ngay  = Day(EndTime)
                e_Thang = Month(EndTime)
                e_Nam   = Year(EndTime)
                e_Gio   = Hour(EndTime)
                e_Phut  = Minute(EndTime)
                e_Giay  = Second(EndTime)
    
                IF   (b_Thang = e_Ngay)  and (b_Nam = e_nam) THEN 
                      e_DTime = e_Ngay&"-"&e_Thang&"-"&e_Nam&" " &e_Gio&":"&e_Phut&":"&e_Giay

                      IF IsDate(e_DTime) THEN
                          Total_time = datediff("'"&TTime&"'",l_BeginTime,e_DTime)' in phút
                      ELSE
                        ' Date time is not format                         
                      END IF
                ELSE
                    'Date time is not format    
                END  IF
             ELSE
                 Total_time = datediff("n",l_BeginTime,l_EndTime) ' in phút                    
             End IF
             'DateTime_Total = CDate(BeginTime)  &" <br />" &CDate(l_EndTime) &"<br />" &Total_time  &"Tháng : "&Thang
             DateTime_Total =Total_time&donvi  
        END IF
    End Function
%>





<%
    Function Obj_val(Tb_name,Col,condition)
        IF Tb_name <> "" And  Col <> "" THEN
            sql_obj = "SELECT "&Col&" FROM "&Tb_name&" "&condition
            set rs_obj = Server.CreateObject("ADODB.RecordSet")
            rs_obj.Open sql_obj,con,1
            IF not rs_obj.EOF THEN
                Obj_val = Trim(rs_obj(Col))
            END IF
        END IF
    End Function
%>
<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function GetMaxId(TableName, FieldNameId, sCondition)
	Dim Max,rsMaxId
	set rsMaxId=server.CreateObject("ADODB.Recordset")
	sql="select Max(" & FieldNameId & ") as MaxId from " & TableName
	if sCondition<>"" then
		sql=sql & " where sCondition"
	end if
	rsMaxId.Open sql, con, 1
	if IsNull(rsMaxId("MaxId")) then
		Max=1
	else
		Max=CLng(rsMaxId("MaxId")) +1
	end if
	rsMaxId.close
	set rsMaxId=nothing
	GetMaxId=Max
End Function
%>

<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getDateServer() 
    '9/6/2016 3:37:52 PM-  MM/DD/YYYY H:I:S
    sqltime = "SELECT GETDATE() AS SDateTime"
    Set rsTime=Server.CreateObject("ADODB.Recordset")
	rsTime.open sqltime,con,1
    IF NOT rsTime.EOF THEN 
        getDateServer = Trim(rsTime("SDateTime"))
    ELSE
       ' getDateServer = Now
    END IF
    rsTime.close
	set rsTime=nothing
End function %>


<%'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function fXungHo(Ten)
	if Ten <> "" then
	arNam	=	Array("Ông","Mr","Văn","Nam","Hưng","Hùng","Quang","Tuấn","Hiếu","Thành","Quân","Dũng","Duy","Tùng","Trí","Sơn","Trưng","Thăng","Thắng","Việt","Trường","Tiến","Toàn","Đức","Cường","Hiệp","Long","Thọ","Khoa","Trọng","Công","Cương","Phong","Kiên","Hưng","Huy","Bộ","Đông","Thọ","Lâm","Mạnh","Xây","Đạt","Hữu","Thịnh","Sĩ","Đình","Tiệp","Tuân","Bá","Tân","Tấn","Đại","Vĩnh","Tính","Dầu","Tý","Tam","Lục","Trung")
	
	arNu	=	Array("Miss","Ms","Bà","Chị","Thị","Thi","Gái","Nữ","Hằng","Hương","Chi","Trang","Thúy","Lan","Ngân","Huyền","Liên","Huệ","Đào","Liễu","Nhung","Trúc","Mai","Nga","Quyên","Lệ","Yến","Nguyệt","Bích","Hoài","Hảo","Thảo","Thơm","Loan","Uyên","Lành","Tươi","Dung","Vy","Hoa","Diệp","Điệp","Ly","Diệu","Thắm","Nhi","Duyên","Thêu","Thủy","Tiên","Kim","Khánh","Phượng","Gấm")
	
	arNamNu	=	Array("Anh","Ánh","Chung","Ngọc","Hải","Giang","Châu","Phương","Thu","Hà","Minh","Yên"," Thương","Vui","Thụy","Sự","Thu","Bình","Hạnh","Xuân","Thuận","Bạn","Thùy","Thanh","Nguyên","Dương","Tú","Châu","Sáu","Bảy","Hai","Tư","Quỳnh")
	
	isNam	=	0
	isNu	=	0
	isNamNu	=	0

	TenTemp=	Ucase(Ten)
	arTen	=	split(TenTemp," ")
	lTen	=	UBound(arTen)
	for k = 0 to lTen
		for t=0 to Ubound(arNam) 
			if StrComp(Ucase(arNam(t)),arTen(k)) =0 then
				isNam	=	isNam	+ 	1
			end if
		next
		
		for t = 0 to UBound(arNu)
			if StrComp(Ucase(arNu(t)),arTen(k)) = 0 then
				isNu	=	isNu	+ 	1
			end if
		next
		for t = 0 to UBound(arNamNu)
			if StrComp(Ucase(arNamNu(t)),arTen(k))=0 then
				isNamNu	=	isNamNu	+ 	1
			end if
		next
	next
	'Response.Write(isNam&" "&isNu&" "&isNamNu)
	if isNam > isNu then
		xungho	=	"anh"
	elseif isNu > isNam then
		xungho	=	"chị"
	elseif isNam = isNu and isNam <> 0 then
		xungho	=	"bạn"
	elseif isNamNu >= 3 then
		xungho	=	"chị"
	else
		xungho	=	"bạn"		
	end if
		
	fXungHo	=	xungho		
	else
		fXungHo=""
	end if
end function
%>



<%Function Uni2NONE(sStr)
	Dim sTemp
	sTemp=Trim(sStr)
	
	'a
	sTemp=Replace(sTemp,"á","a")
	sTemp=Replace(sTemp,"à","a")
	sTemp=Replace(sTemp,"ả","a")
	sTemp=Replace(sTemp,"ã","a")
	sTemp=Replace(sTemp,"ạ","a")
	
	'ă
	sTemp=Replace(sTemp,"ă","a")
	sTemp=Replace(sTemp,"ắ","a")
	sTemp=Replace(sTemp,"ằ","a")
	sTemp=Replace(sTemp,"ẳ","a")
	sTemp=Replace(sTemp,"ẵ","a")
	sTemp=Replace(sTemp,"ặ","a")
	
	'â
	sTemp=Replace(sTemp,"â","a")
	sTemp=Replace(sTemp,"ấ","a")
	sTemp=Replace(sTemp,"ầ","a")
	sTemp=Replace(sTemp,"ẩ","a")
	sTemp=Replace(sTemp,"ẫ","a")
	sTemp=Replace(sTemp,"ậ","a")
	
	'đ
	sTemp=Replace(sTemp,"đ","d")
	
	'e
	sTemp=Replace(sTemp,"é","e")
	sTemp=Replace(sTemp,"è","e")
	sTemp=Replace(sTemp,"ẻ","e")
	sTemp=Replace(sTemp,"ẽ","e")
	sTemp=Replace(sTemp,"ẹ","e")
	
	'ê
	sTemp=Replace(sTemp,"ê","e")
	sTemp=Replace(sTemp,"ế","e")
	sTemp=Replace(sTemp,"ề","e")
	sTemp=Replace(sTemp,"ể","e")
	sTemp=Replace(sTemp,"ễ","e")
	sTemp=Replace(sTemp,"ệ","e")
	
	'i
	sTemp=Replace(sTemp,"í","i")
	sTemp=Replace(sTemp,"ì","i")
	sTemp=Replace(sTemp,"ỉ","i")
	sTemp=Replace(sTemp,"ĩ","i")
	sTemp=Replace(sTemp,"ị","i")
	
	'o
	sTemp=Replace(sTemp,"ó","o")
	sTemp=Replace(sTemp,"ò","o")
	sTemp=Replace(sTemp,"ỏ","o")
	sTemp=Replace(sTemp,"õ","o")
	sTemp=Replace(sTemp,"ọ","o")
	
	'ô
	sTemp=Replace(sTemp,"ô","o")
	sTemp=Replace(sTemp,"ố","o")
	sTemp=Replace(sTemp,"ồ","o")
	sTemp=Replace(sTemp,"ổ","o")
	sTemp=Replace(sTemp,"ỗ","o")
	sTemp=Replace(sTemp,"ộ","o")
	
	'ơ
	sTemp=Replace(sTemp,"ơ","o")
	sTemp=Replace(sTemp,"ớ","o")
	sTemp=Replace(sTemp,"ờ","o")
	sTemp=Replace(sTemp,"ở","o")
	sTemp=Replace(sTemp,"ỡ","o")
	sTemp=Replace(sTemp,"ợ","o")
	
	'u
	sTemp=Replace(sTemp,"ú","u")
	sTemp=Replace(sTemp,"ù","u")
	sTemp=Replace(sTemp,"ủ","u")
	sTemp=Replace(sTemp,"ũ","u")
	sTemp=Replace(sTemp,"ụ","u")
	
	'ư
	sTemp=Replace(sTemp,"ư","u")
	sTemp=Replace(sTemp,"ứ","u")
	sTemp=Replace(sTemp,"ừ","u")
	sTemp=Replace(sTemp,"ử","u")
	sTemp=Replace(sTemp,"ữ","u")
	sTemp=Replace(sTemp,"ự","u")
	
	'y
	sTemp=Replace(sTemp,"ý","y")
	sTemp=Replace(sTemp,"ỳ","y")
	sTemp=Replace(sTemp,"ỷ","y")
	sTemp=Replace(sTemp,"ỹ","y")
	sTemp=Replace(sTemp,"ỵ","y")
'---------------------------------Chữ hoa-------------------------------------------------
	'A
	sTemp=Replace(sTemp,"Á","A")
	sTemp=Replace(sTemp,"À","A")
	sTemp=Replace(sTemp,"Ả","A")
	sTemp=Replace(sTemp,"Ã","A")
	sTemp=Replace(sTemp,"Ạ","A")
	
	'Ă
	sTemp=Replace(sTemp,"Ă","A")
	sTemp=Replace(sTemp,"Ắ","A")
	sTemp=Replace(sTemp,"Ằ","A")
	sTemp=Replace(sTemp,"Ẳ","A")
	sTemp=Replace(sTemp,"Ẵ","A")
	sTemp=Replace(sTemp,"Ặ","A")
	
	'Â
	sTemp=Replace(sTemp,"Â","A")
	sTemp=Replace(sTemp,"Ấ","A")
	sTemp=Replace(sTemp,"Ầ","A")
	sTemp=Replace(sTemp,"Ẩ","A")
	sTemp=Replace(sTemp,"Ẫ","A")
	sTemp=Replace(sTemp,"Ậ","A")
	
	'Đ
	sTemp=Replace(sTemp,"Đ","D")
	
	'E
	sTemp=Replace(sTemp,"É","E")
	sTemp=Replace(sTemp,"È","E")
	sTemp=Replace(sTemp,"Ẻ","E")
	sTemp=Replace(sTemp,"Ẽ","E")
	sTemp=Replace(sTemp,"Ẹ","E")
	
	'Ê
	sTemp=Replace(sTemp,"Ê","E")
	sTemp=Replace(sTemp,"Ế","E")
	sTemp=Replace(sTemp,"Ề","E")
	sTemp=Replace(sTemp,"Ể","E")
	sTemp=Replace(sTemp,"Ễ","E")
	sTemp=Replace(sTemp,"Ệ","E")
	
	'I
	sTemp=Replace(sTemp,"Í","I")
	sTemp=Replace(sTemp,"Ì","I")
	sTemp=Replace(sTemp,"Ỉ","I")
	sTemp=Replace(sTemp,"Ĩ","I")
	sTemp=Replace(sTemp,"Ị","I")
	
	'O
	sTemp=Replace(sTemp,"Ó","O")
	sTemp=Replace(sTemp,"Ò","O")
	sTemp=Replace(sTemp,"Ỏ","O")
	sTemp=Replace(sTemp,"Õ","O")
	sTemp=Replace(sTemp,"Ọ","O")
	
	'Ô
	sTemp=Replace(sTemp,"Ô","O")
	sTemp=Replace(sTemp,"Ố","O")
	sTemp=Replace(sTemp,"Ồ","O")
	sTemp=Replace(sTemp,"Ổ","O")
	sTemp=Replace(sTemp,"Ỗ","O")
	sTemp=Replace(sTemp,"Ộ","O")
	
	'Ơ
	sTemp=Replace(sTemp,"Ơ","O")
	sTemp=Replace(sTemp,"Ớ","O")
	sTemp=Replace(sTemp,"Ờ","O")
	sTemp=Replace(sTemp,"Ở","O")
	sTemp=Replace(sTemp,"Ỡ","O")
	sTemp=Replace(sTemp,"Ợ","O")
	
	''U
	sTemp=Replace(sTemp,"Ú","U")
	sTemp=Replace(sTemp,"Ù","U")
	sTemp=Replace(sTemp,"Ủ","U")
	sTemp=Replace(sTemp,"Ũ","U")
	sTemp=Replace(sTemp,"Ụ","U")
	
	''Ư
	sTemp=Replace(sTemp,"Ư","U")
	sTemp=Replace(sTemp,"Ứ","U")
	sTemp=Replace(sTemp,"Ừ","U")
	sTemp=Replace(sTemp,"Ử","U")
	sTemp=Replace(sTemp,"Ữ","U")
	sTemp=Replace(sTemp,"Ự","U")
	
	''Y
	sTemp=Replace(sTemp,"Ý","Y")
	sTemp=Replace(sTemp,"Ỳ","Y")
	sTemp=Replace(sTemp,"Ỷ","Y")
	sTemp=Replace(sTemp,"Ỹ","Y")
	sTemp=Replace(sTemp,"Ỵ","Y")

	'ký tự thừa
	sTemp=Replace(sTemp,"/","")
	sTemp=Replace(sTemp,"\","")
	sTemp=Replace(sTemp,",","")
	sTemp=Replace(sTemp,"&","")
	sTemp=Replace(sTemp,"$","")
	sTemp=Replace(sTemp,"~","")
	sTemp=Replace(sTemp,"*","")
	sTemp=Replace(sTemp,"(","")
	sTemp=Replace(sTemp,")","")
	sTemp=Replace(sTemp,"{","")	
	sTemp=Replace(sTemp,"}","")
	sTemp=Replace(sTemp,"|","")
	sTemp=Replace(sTemp,"'","''")
	sTemp=Replace(sTemp,"  ","")
    sTemp=replace(sTemp,"?","")
    sTemp=replace(sTemp,"%","")
    sTemp=replace(sTemp,":","")	
    sTemp=replace(sTemp,"+","")	
    sTemp=replace(sTemp,"--","-")	
    Uni2NONE=sTemp
End Function
%>