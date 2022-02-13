
<%Function maxx(a,b)
	if a > b then
		maxx = a
	else
		maxx = b
	end if
End Function%>

<%Function minn(a,b)
	if a < b then
		minn = a
	else
		minn = b
	end if
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