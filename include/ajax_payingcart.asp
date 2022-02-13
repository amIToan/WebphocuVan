<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/JSON_2.0.4.asp" -->


<script language="javascript" runat="server" src='/javascript/json2.js'></script>
<%  
    keyWord = Trim(Request.form("keyword")) 
    Select Case keyWord
    Case "cartlists"
        Call cartLists()
    Case else
        Call payingcart()
    End Select

%>

<%
    Sub cartLists()
        discountedKey = Trim(Request.Form("discountedKey"))
        sql_Voucher = "Select * from Voucher as Vou where Vou.Voucher = N'"&discountedKey&"'"
        Set rsVoucher = Server.CreateObject("ADODB.Recordset")
        rsVoucher.open sql_Voucher,con,1
        If not rsVoucher.eof then
            discountVal = rsVoucher("value")
            Response.ContentType = "application/json"
            Response.Write("{ ""status"":""1"",""discountVal"":"""&discountVal&""" }")
        else
            Response.ContentType = "application/json"
            Response.Write("{ ""status"":""0""}")
        End if
    End Sub 
%>
<% 
    Sub payingcart()
    name_Cus = Request.Form("fullName")
    email_Cus = Request.Form("emailUSer")
    phone_Cus = Request.Form("phoneNumber")
    provinceid_Cus = Request.Form("Province")
    districtid_Cus = Request.Form("District")
    wardid_Cus = Request.Form("Ward")
    address_Cus = Request.Form("Address")
    if  Request.Form("postMethod") <> "" then
        method_Cus = Request.Form("postMethod")
    elseif Request.Form("interbanking") then 
        method_Cus = Request.Form("interbanking")
    End if   
    Dim myJSON
    myJSON = Request.Form("strProduct") 
    Set rsCustomer = Server.CreateObject("ADODB.Recordset")
    sqlcheck ="Select * From Customers Where (Phone=N'"&phone_Cus&"' and email='"&email_Cus&"')"
    rsCustomer.open sqlcheck,con,1
    if not rsCustomer.EOF then
        sql_update="Update Customers set Count="&(Clng(rsCustomer("count")) + 1)&" where (Phone=N'"&phone_Cus&"' and email='"&email_Cus&"')"
        con.execute(sql_update)
        on error resume next  
        if err<>0 then 
            Response.ContentType = "application/json"
            Response.Write("{ ""status"":""0""}")
        else
             sql_Order= " Insert into Orders( CustomerID, OrderDate, ShipAddress, ShipWard, ShipCity, ShipRegion, ShipVia)"&_
                        "VALUES ('"&rsCustomer("CustomerID")&"',FORMAT (getdate(), 'dd/MM/yyyy, hh:mm'),N'"&address_Cus&"','"&wardid_Cus&"','"&districtid_Cus&"','"&provinceid_Cus&"', '"&method_Cus&"')"
            con.execute(sql_Order)
            if err<>0 then 
                Response.ContentType = "application/json"
                Response.Write("{ ""status"":""0""}")
            else
                sql_getOrderID = "SELECT TOP 1 * FROM Orders ORDER BY OrderID DESC "
                set getOrderID = Server.CreateObject("ADODB.Recordset")
                getOrderID.open sql_getOrderID,con,1
                    Set myJSON = JSON.parse(myJSON) 
                        For each Item in myJSON
                            sqlOrderDetails = " Insert into OrderDetails( OrderID, ProductID, UnitPrice, Quantity,imglinks)"&_
                                                "VALUES ('"&getOrderID("OrderID")&"',N'"&Item.id&"','"&Item.price&"','"&Item.count&"',N'"&Item.img&"')"
                            con.execute(sqlOrderDetails)
                        Next
                            if err<>0 then 
                                Response.ContentType = "application/json"
                                Response.Write("{ ""status"":""0""}")
                            else
                                Response.ContentType = "application/json"
                                Response.Write("{ ""status"":""1""}")
                            end IF
                Set getOrderID = nothing
            End if
        end if
    else
    sqlInsert = "INSERT INTO Customers (contactName, Phone, email, region, City, ward, Address , count) VALUES (N'"&name_Cus&"',N'"&phone_Cus&"',N'"&email_Cus&"',N'"&provinceid_Cus&"',N'"&districtid_Cus&"',N'"&wardid_Cus&"',N'"&address_Cus&"','1')"
        con.execute(sqlInsert)
        on error resume next  
        if err<>0 then 
            Response.ContentType = "application/json"
            Response.Write("{ ""status"":""0""}")
        else
             Set rsOrderdetails = Server.CreateObject("ADODB.Recordset")
            sql_Order= " Insert into Orders( CustomerID, OrderDate, ShipAddress, ShipWard, ShipCity, ShipRegion, ShipVia)"&_
                        "VALUES ('"&rsCustomer("CustomerID")&"',FORMAT (getdate(), 'dd/MM/yyyy, hh:mm'),N'"&address_Cus&"','"&wardid_Cus&"','"&districtid_Cus&"','"&provinceid_Cus&"', '"&method_Cus&"')"
            con.execute(sql_Order)
            if err<>0 then 
                Response.ContentType = "application/json"
                Response.Write("{ ""status"":""0""}")
            else
                sql_getOrderID = "SELECT TOP 1 * FROM Orders ORDER BY OrderID DESC "
                set getOrderID = Server.CreateObject("ADODB.Recordset")
                getOrderID.open sql_getOrderID,con,1
                    Set myJSON = JSON.parse(myJSON) 
                        For each Item in myJSON
                            sqlOrderDetails = " Insert into OrderDetails( OrderID, ProductID, UnitPrice, Quantity, imglinks)"&_
                                                "VALUES ('"&getOrderID("OrderID")&"',N'"&Item.id&"','"&Item.price&"','"&Item.count&"',N'"&Item.img&"')"
                            con.execute(sqlOrderDetails)
                        Next
                            if err<>0 then 
                                Response.ContentType = "application/json"
                                Response.Write("{ ""status"":""0""}")
                            else
                                Response.ContentType = "application/json"
                                Response.Write("{ ""status"":""1"",""order"":"""&getOrderID("OrderID")&"""}")
                            end IF
                Set getOrderID = nothing
            End if
        end if
    End if
 Set rsCustomer = nothing
 End Sub
%>
