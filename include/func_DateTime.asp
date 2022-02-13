<%Function GetFullDateVn(theDate)
	ngay = CInt(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = CInt(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = CInt(year(theDate))
	GetFullDateVn="Ngày "&ngay & " Tháng " & thang & " Năm " & nam
End Function%>

<%Function GetFullDate(theDate)
	ngay = CInt(day(theDate))
	if ngay<10 then
		ngay="0" & ngay
	end if
    thang = CInt(month(theDate))
    if thang<10 then
		thang="0" & thang
	end if
    nam = CInt(year(theDate))
    
	GetFullDate=ngay & "/" & thang & "/" & nam
End Function%>

<%Function GetFullTime(theDate)
	gio=Cint(Hour(theDate))
	if gio<10 then
		gio="0" & gio
	end if
	phut=Cint(Minute(theDate))
	if phut<10 then
		phut="0" & phut
	end if

	GetFullTime=gio & "h" & phut & """&nbsp;"
End Function%>

<%Function GetNameOfWeekDay(theDate)
	Select case Weekday(thedate)
		case 1
			strWeekDay="Ch&#7911; nh&#7853;t"
		case 2
			strWeekDay="Th&#7913; hai"
		case 3
			strWeekDay="Th&#7913; ba"
		case 4
			strWeekDay="Th&#7913; t&#432;"
		case 5
			strWeekDay="Th&#7913; n&#259;m"
		case 6
			strWeekDay="Th&#7913; s&#225;u"
		case 7
			strWeekDay="Th&#7913; b&#7849;y"
	End select
	GetNameOfWeekDay=strWeekDay
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
		for i=YearStart to Year(now())+1 
			Response.write "<option value=""" & i & """"
			if i=Clng(YearSelected) then
				response.write " selected"
			end if
			response.write ">" & i & "</option>" & vbNewline
		next
	response.write "</select>" & vbNewline
End Sub%>


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