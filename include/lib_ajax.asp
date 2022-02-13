﻿<%
    Function func_uri(cateid,newsid,title)
        link = ""
        if cateid <> "" then link = link & "/"&cateid
        if newsid <> "" then link = link & "/"&newsid
        if title <> "" then link = "/"& Replace(Uni2NONE(lcase(title))," ","-")&link&".html"
        if link <> "" then func_uri = link else func_uri = "#"
    End Function 
%>


<%
Function Uni2NONE(sStr)
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
    Uni2NONE=sTemp
End Function
'==========================================================
Function Uni2NONE1(sStr)
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
	sTemp=Replace(sTemp,":","")
	sTemp=Replace(sTemp,"/","")
	sTemp=Replace(sTemp,"\","")
	sTemp=Replace(sTemp,",","")
	sTemp=Replace(sTemp,"&","")
	sTemp=Replace(sTemp,"$","")
	sTemp=Replace(sTemp,"~","")
	sTemp=Replace(sTemp,"*","")
	sTemp=Replace(sTemp,"#","")
	sTemp=Replace(sTemp,"(","")
	sTemp=Replace(sTemp,")","")
	sTemp=Replace(sTemp,"{","")	
	sTemp=Replace(sTemp,"}","")
	sTemp=Replace(sTemp,"|","")
	sTemp=Replace(sTemp,"'","''")
	sTemp=Replace(sTemp,"  ","-")
	sTemp=Replace(sTemp," ","-")	
    	
    sTemp=replace(sTemp,"?","")
    sTemp=replace(sTemp,"%"," phan-tram")
	
    Uni2NONE1=LCase(sTemp)
End Function
%>



<%
function Dis_str_money(moneys)
    if not isNumeric(moneys) then
        Dis_str_money = 0
        exit function
    end if
	moneys	=	round(Cdbl(moneys))
	str_dao=Daochuoi(moneys)
	lenn=Len(str_dao)
	k3=1
	money_re=""
	for bieni=1 to lenn
		money_re=money_re+Mid(str_dao,bieni,1)
		if k3=3 and bieni< lenn then
			money_re=money_re+"."
			k3=1
		else
			k3=k3+1
		end if
	next
	Dis_str_money=Daochuoi(money_re)
end function

function Daochuoi(str_money)
	dim str_temp
	k_lengh=len(str_money)
	for bieni=1 to k_lengh
		str_temp=MID(str_money,bieni,1)+str_temp
		'Response.write("i="& i &" str_money "& str_money)
	next
	Daochuoi=str_temp
end function
%>