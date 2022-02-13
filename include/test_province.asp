<!--#include virtual="/include/config.asp" -->
<!--#include virtual="include/JSON_2.0.4.asp" -->
 <%
    keyValue = Request.QueryString("keyValue")
    ProvinceID = Request.QueryString("ProvinceID")
    DistrictID = Clng(Request.QueryString("DistrictID"))
    Select Case keyValue
    Case "province"
        Call Get_district(ProvinceID)
    Case "district"
        Call Get_ward(DistrictID)
    Case else
        response.write("No value!!!")
    End Select
 %>
<% 
'get-district
Sub Get_district(ProvinceID)    
    
    html=""
    status= false
    sqlNews = "Select DistrictID,NameDistrict From Province_district Where ProvinceID="&ProvinceID   
    Set rsNews=Server.CreateObject("ADODB.Recordset")
    'Response.write(sqlNews)
    rsNews.open sqlNews,con,3 
    If not rsNews.EOF Then
        status=True
        'html=html&"<option value='0'>Chọn quận / huyện</option>"
        Do while not rsNews.eof
            html=html&"<option value='"&rsNews("DistrictID")&"'>"&rsNews("NameDistrict")&"</option>"
        rsNews.movenext
    Loop 
    End If 'end if rsNews.EOF
    Set rsNews=nothing

    'xuất chuỗi json ra ngoài màn hình
    Set project = jsObject()
    Set project("error") = jsObject()
    project("error")("status") = status
    project("data") = html
    project.Flush
End Sub %>

<% 
'get-ward
Sub Get_ward(DistrictID)    
    
    html=""
    status= false
    sqlNews = "Select wardID,wardName From Province_ward Where DistrictID="&DistrictID   
    Set rsNews=Server.CreateObject("ADODB.Recordset")
    'Response.write(sqlNews)
    rsNews.open sqlNews,con,3 
    If not rsNews.EOF Then
        status=True
        'html=html&"<option value='0'>Chọn phường xã</option>"
        Do while not rsNews.eof
            html=html&"<option value='"&rsNews("wardID")&"'>"&rsNews("wardName")&"</option>"
        rsNews.movenext
    Loop 
    End If 'end if rsNews.EOF
    Set rsNews=nothing

    'xuất chuỗi json ra ngoài màn hình
    Set project = jsObject()
    Set project("error") = jsObject()
    project("error")("status") = status
    project("data") = html
    project.Flush
End Sub %>

<% 
'get-street
Sub Get_street(wardID)    
    
    html=""
    status= false
    sqlNews = "Select streetID,streetName From Province_street Where wardID="&wardID   
    Set rsNews=Server.CreateObject("ADODB.Recordset")
    'Response.write(sqlNews)
    rsNews.open sqlNews,con,3 
    If not rsNews.EOF Then
        status=True
        Do while not rsNews.eof
            html=html&"<option value='"&rsNews("streetID")&"'>"&rsNews("streetName")&"</option>"
        rsNews.movenext
    Loop 
    End If 'end if rsNews.EOF
    Set rsNews=nothing

    'xuất chuỗi json ra ngoài màn hình
    Set project = jsObject()
    Set project("error") = jsObject()
    project("error")("status") = status
    project("data") = html
    project.Flush
End Sub %>
