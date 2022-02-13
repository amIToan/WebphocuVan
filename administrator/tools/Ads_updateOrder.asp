<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_Ads.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
If Request.Form("Submit") = "Update Order" Then
CatID=Request.Form("CatID")
set rsCommon=server.CreateObject("ADODB.Recordset")
	sql="SELECT a.Ads_id, a.Ads_Title, a.Ads_width, a.Ads_height, a.StatusId, a.Ads_Creator, " &_ 
		"		a.Ads_CreationDate, a.Ads_LastEditor, a.Ads_LastEdited, a.Ads_Position, " &_
        "		a.Ads_ImagesPath,a.Ads_orderSub, d.CategoryId, d.Ads_OnlineChildren, d.Ads_Order " &_
		"FROM   Ads a INNER JOIN" &_
        "			AdsDistribution d ON a.Ads_id = d.Ads_id "

		''Li?t kê trong m?t chuyên m?c
			LstCat=GetListParentCat(CatId)
			ArrLstCat=Split(" " & Trim(LstCat))
			
    	sql=sql & " WHERE"
		sql=sql & " (d.CategoryId=0 or d.CategoryId=" & CatId
			For i=1 to UBound(ArrLstCat)
				if IsNumeric(ArrLstCat(i)) then
					sql=sql & " or (d.CategoryId=" & Clng(ArrLstCat(i)) & " and d.Ads_OnlineChildren=1)"
				end if
			Next
			sql=sql & ")"
	sql=sql & " ORDER BY d.Ads_Order desc"

	rsCommon.CursorType = 2
	rsCommon.LockType = 3
	rsCommon.Open sql, con
	Do While NOT rsCommon.EOF
	
		rsCommon.Fields("Ads_order") = Clng(Request.Form("AdsOrder" & rsCommon("Ads_ID")))		
		rsCommon.Update
	rsCommon.MoveNext
	Loop
	rsCommon.Close
End If
Set rsCommon = Nothing
con.Close
Set con = Nothing
Response.Write("OK")
Response.End
Response.Redirect "configure_category.asp"
%>