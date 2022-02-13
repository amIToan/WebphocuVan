
<%Function ReplaceTextToHTML(Str)
	Dim strTmp
	strTmp=Trim(Str)
	strTmp=Replace(strTmp,chr(13) & chr(10),"<br>")
	strTmp=Replace(strTmp,"'","''")
	strTmp=Replace(strTmp,"""","&quot;")
	ReplaceTextToHTML=strTmp
End Function%>

<%Function ReplaceHTMLToText(Str)
	Dim strTmp
	strTmp=Trim(Str)
	strTmp=Replace(strTmp,"<br>",chr(13) & chr(10))
	strTmp=Replace(strTmp,"''","'")
	ReplaceHTMLToText=strTmp
End Function%>

<%Function GetNumeric(sValue,DefaultValue)
	Dim intValue
	if not IsNumeric(sValue) or trim(sValue)="" then
		intValue=DefaultValue
	else
		intValue=Clng(sValue)
	end if
	GetNumeric=intValue
End Function%>