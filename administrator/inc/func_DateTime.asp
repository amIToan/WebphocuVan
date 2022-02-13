<%Function GetFullDate(theDate,lang)
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
<%Function GetFullDateTime(theDate)
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

<%Sub List_Date(DateSelected,DateTitle)
	response.write "<select name=""Ngay"">" & vbNewline &_
		"<option value=""0"">" & DateTitle & "</option>" & vbNewline
		for i=1 to 31 
			Response.write "<option value=""" & i & """"
			if i=Clng(DateSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>
<%Sub List_Month(MonthSelected,MonthTitle)
	response.write "<select name=""Thang"">" & vbNewline &_
		"<option value=""0"">" & MonthTitle & "</option>" & vbNewline
		for i=1 to 12 
			Response.write "<option value=""" & i & """"
			if i=Clng(MonthSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>
<%Sub List_Year(YearSelected,YearTitle,YearStart)
	response.write "<select name=""Nam"">" & vbNewline &_
		"<option value=""0"">" & YearTitle & "</option>" & vbNewline
		for i=YearStart to Year(now()+1) 
			Response.write "<option value=""" & i & """"
			if i=Clng(YearSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>

<%Sub List_Date_WithName(DateSelected,DateTitle,DateName)
	response.write "<select name=""" & DateName & """>" & vbNewline &_
		"<option value=""0"">" & DateTitle & "</option>" & vbNewline
		for i=1 to 31 
			Response.write "<option value=""" & i & """"
			if i=Clng(DateSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>
<%Sub List_Month_WithName(MonthSelected,MonthTitle,MonthName)
	response.write "<select name=""" & MonthName & """>" & vbNewline &_
		"<option value=""0"">" & MonthTitle & "</option>" & vbNewline
		for i=1 to 12 
			Response.write "<option value=""" & i & """"
			if i=Clng(MonthSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>
<%Sub List_Year_WithName(YearSelected,YearTitle,YearStart,YearName)
	response.write "<select name=""" & YearName & """>" & vbNewline &_
		"<option value=""0"">" & YearTitle & "</option>" & vbNewline
		for i=YearStart to Year(now()) +1
			Response.write "<option value=""" & i & """"
			if i=Clng(YearSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>




<%'VIETSOFT 03/08/2015 %>

<%
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


<%
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



<%
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
'  '+ - time
'  Function DateTime_(BeginTime,EndTime)     
'      IF BeginTime <> "" and  EndTime <> ""  and IsDate (BeginTime) and IsDate(EndTime) THEN  
'           l_BeginTime = Replace(BeginTime, "/", "-")
'           l_EndTime   = Replace(EndTime, "/", "-")
'           Total_time = datediff("n",l_BeginTime,l_EndTime) ' in phút
'           IF Total_time  < 0 THEN 
'              b_Ngay  = Day(BeginTime)
'              b_Thang = Month(BeginTime)
'              b_Nam   = Year(BeginTime)
'              b_Gio   = Hour(BeginTime)
'              b_Phut  = Minute(BeginTime)
'              b_Giay  = Second(BeginTime)
'
'              e_Ngay  = Day(EndTime)
'              e_Thang = Month(EndTime)
'              e_Nam   = Year(EndTime)
'              e_Gio   = Hour(EndTime)
'              e_Phut  = Minute(EndTime)
'              e_Giay  = Second(EndTime)
'
'              IF   (b_Thang = e_Ngay)  and (b_Nam = e_nam) THEN 
'                    e_DTime = e_Ngay&"-"&e_Thang&"-"&e_Nam&" " &e_Gio&":"&e_Phut&":"&e_Giay
'
'                    IF IsDate(e_DTime) THEN
'                        Total_time = datediff("n",l_BeginTime,e_DTime)' in phút
'                    ELSE
'                      ' Date time is not format                         
'                    END IF
'              ELSE
'                  'Date time is not format    
'              END  IF
'           ELSE
'               Total_time = datediff("n",l_BeginTime,l_EndTime) ' in phút                    
'           End IF
'           'DateTime_Total = CDate(BeginTime)  &" <br />" &CDate(l_EndTime) &"<br />" &Total_time  &"Tháng : "&Thang
'           DateTime_Total =Total_time  
'      END IF
'  End Function
%>