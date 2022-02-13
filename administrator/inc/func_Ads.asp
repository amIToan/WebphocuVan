<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Sub List_Ads_Position(AdsSelect,Ads_Position_Title)%>
		<select name="Ads_Position" id="Ads_Position">
			<option value="-1"><%=Ads_Position_Title%></option>
			<option value="8"<%if Clng(AdsSelect)=8 then%> selected <%end if%>>Top center banner(1349x351)</option>
			<option value="7"<%if Clng(AdsSelect)=7 then%> selected <%end if%>>Banner quảng cáo nhỏ</option>
			<option value="6"<%if Clng(AdsSelect)=7 then%> selected <%end if%>>Menu Phải (Dưới 386xAuto)</option>
			<option value="0"<%if Clng(AdsSelect)=0 then%> selected <%end if%>>Trên nội dung</option>
			<option value="2"<%if Clng(AdsSelect)=0 then%> selected <%end if%>>Dưới nội dung</option>
	  		<option value="1"<%if Clng(AdsSelect)=1 then%> selected <%end if%>>Dưới Nội dung chuyên mục</option>
			<option value="9"<%if Clng(AdsSelect)=9 then%> selected <%end if%>>Dưới index</option>
        </select>
<%End sub%>
<%Function Get_Ads_Position_Name(Ads_Position)
	Select case Ads_Position
		case 8
			Get_Ads_Position_Name="Top center banner[1349x351]"	
		case 7
			Get_Ads_Position_Name="Banner quảng cáo nhỏ"
        case 6
			Get_Ads_Position_Name="Dưới Menu Phải [386xAuto]"
		case 0
			Get_Ads_Position_Name="Trên Nội dung chi tiết [Auto]"
        case 2
			Get_Ads_Position_Name="Dưới Nội dung chi tiết [Auto]"
		case 1
			Get_Ads_Position_Name="Dưới Nội dung chuyên mục [Auto]"
		case 9
			Get_Ads_Position_Name="Dưới index [Auto]"

		case else
			Get_Ads_Position_Name=""
	End select
End Function%>
<%Sub List_Ads_StatusId(AdsSelect,Ads_StatusId_Title)%>
		<select name="StatusId" id="StatusId">
			<option value="-1"><%=Ads_StatusId_Title%></option>
			<option value="apap"<%if AdsSelect="apap" then%> selected<%end if%>>Đưa lên mạng</option>
			<option value="eded"<%if AdsSelect="eded" then%> selected<%end if%>>Không lên mạng</option>
        </select>
<%End sub%>
<%Function Get_Ads_StatusId_Name(StatusId)
	Select case StatusId
		case "eded"
			Get_Ads_StatusId_Name="Không lên mạng"
		case "apap"
			Get_Ads_StatusId_Name="Đưa lên mạng"
		case else
			Get_Ads_StatusId_Name=""
	End select
End Function%>

<%Function CheckOnlineChildrenInList(CatId,LstOnlineChildren)
	if Instr(" " & LstOnlineChildren & " "," " & CatId & "1 ")>0 then
		CheckOnlineChildrenInList=true
	else
		CheckOnlineChildrenInList=false
	end if
End Function%>