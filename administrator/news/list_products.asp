<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link href="/stylesheets/w3style.css" rel="stylesheet" />
    <link href="/interfaces/liberary/bootstrap4/css/bootstrap.min.css" rel="stylesheet" />
    <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/jquery-3.5.1.min.js"></script>
    <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/bootstrap.min.js"></script>
</head>
<body>
    <div class="container-fluid p-0">
        <%
            Call header() 
        %>
    </div>
    <%
        sqllistOrder = " Select cus.ContactName, cus.phone, cus.email, o.OrderID, o.OrderDate, o.Shipvia, o.ShipAddress, pw.wardName, pd.NameDistrict, Province.NameProvince from Orders as o inner join Customers as cus on o.CustomerID = cus.CustomerID"&_
                                   " inner join Province on Province.ProvinceID = o.ShipRegion"&_
                                   " inner join Province_district as pd on pd.DistrictID = o.ShipCity"&_
                                   " inner join Province_ward as pw on pw.wardID = o.ShipWard"
        action          =   Request.Form("action")
        keyword         =   Trim(Request.form("keyword"))
        if keyword <>"" then
            if IsNumeric(keyword) then
                isnumber = Clng(keyword)
            End if
            Session("key") = keyword
        else
            Session("key") = ""
        end if
                    If action="Search" Then 
                        if isnumber <> "" then 'if l?? m?? tin
                            sqllistOrder= sqllistOrder +" where o.orderID="&isnumber
                        end if
                    End If
        sqllistOrder= sqllistOrder + " ORDER BY o.OrderID DESC"
    %>
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
        <div class="col-md-10">
            <div class="w3-container">
                <div class="form-horizontal w3-section">
                    <form method="post" action="list_products.asp">
                            <div class="form-group" style="display:flex; align-items: center">
                                <div class="col-md-3">
                                    <h4>T??m ki???m theo ????n h??ng: </h4>
                                </div>
                                <div class="col-md-9 w3-right-align" style="display:flex; align-items: center; justify-content: flex-end">
                                    <input name="keyword" type="text"  value="<%=Session("key")%>" placeholder="Ti??u ?????" class="form-control" >
                                    <input type="submit" name="ButtonSearch" id="ButtonSearch" class="btn btn-primary" value="T??m ki???m" />
                                </div>
                            </div>
                            <input type="hidden" name="action" value="Search">
                    </form>
                </div>
                <div class="w3-responsive w3-section">
                        <table class="w3-table-all w3-hoverable w3-striped w3-center" >
                            <thead>
                                    <tr class="w3-blue" >
                                        <th>STT</th>
                                        <th >T??n kh??ch h??ng</th>
                                        <th>S??T</th>
                                        <th>Email</th>
                                        <th>????n h??ng</th>
                                        <th>Ng??y Order</th>
                                        <th>T??nh tr???ng</th>
                                        <th>?????a ch???</th>
                                        <th colspan="2">Ch???c n??ng</th>
                                    </tr>
                            </thead>
                            <!-- M??? k???t n???i sql-->
                            <%
                    Set rs = Server.CreateObject("ADODB.RecordSet")
                    rs.open sqllistOrder,con,1
                    If not rs.eof Then
                        stt = 1
                        rs.PageSize = 20
                        pagecount = rs.PageCount
                        pagination = 4
                        if request.QueryString("page") <> ""then
                            page = Clng(request.QueryString("page"))
                        else
                            page = 1
                        End if
                        rs.AbsolutePage = CLng(page)
                        Do While not rs.eof and j < rs.PageSize
                                Name                  =   Trim(rs("ContactName"))
                                mobile                =   Trim(rs("phone"))
                                email                 =   Trim(rs("email"))
                                payingStatus          =   Trim(rs("ShipVia"))
                                orderID               =   Trim(rs("OrderID"))
                                orderDate             =   Trim(rs("OrderDate"))
                                Address                = Trim(rs("ShipAddress"))
                                wardName              =   Trim(rs("wardName"))
                                nameDistrict          =   Trim(rs("NameDistrict"))
                                NameProvince          =   Trim(rs("NameProvince"))
                    %>
                    <tr>
                        <td><% = stt%></td>
                        <td><% = Name%></td>
                        <td><% = mobile%></td>
                        <td><% = email%></td>
                        <td><% = orderID%></td>
                        <td><% = orderDate%></td>
                        <td><%
                            If payingStatus = 0 then
                                response.write("Thanh to??n khi nh???n")
                            else
                                response.write("Thanh to??n tr?????c")
                            END if
                         %></td>
                         <td style="width: 300px">
                             <% response.write(Address&", "&wardName&", "&nameDistrict&", "&NameProvince) %>
                         </td>
                        <td>
                            <button class="btn w3-blue" data-toggle="modal" data-target="#myModal<% = orderID%>">Chi ti???t</button>
                            <button class="btn btn-danger">X??a</button>
                            <!-- The Modal -->
                            <div class="modal fade" id="myModal<% = orderID%>">
                                <div class="modal-dialog modal-lg">
                                    <div class="modal-content">

                                        <!-- Modal Header -->
                                        <div class="modal-header">
                                            <h4 class="modal-title">Th??ng tin chi ti???t ????n h??ng</h4>
                                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                                        </div>

                                        <!-- Modal body -->
                                        <%
                                        SqlorderDetails = " select od.ProductID, od.Quantity, od.imglinks , od.UnitPrice, n.Title, n.Price from OrderDetails as od inner join News as N on od.ProductID = n.NewsID where od.OrderID="&orderID 
                                        Set rsDetails = Server.CreateObject("ADODB.RecordSet")
                                        rsDetails.open SqlorderDetails,con,1
                                        IF not rsDetails.eof then
                                        %>
                                        <div class="modal-body">
                                            <table class="w3-table-all w3-centered w3-hoverable" >
                                                <thead >
                                                    <tr style="background: #3f7e3b">
                                                        <th>M?? t??? </th>
                                                        <th> S???n ph???m</th>
                                                        <th> Gi?? </th>
                                                        <th> S??? l?????ng</th>
                                                        <th> T???ng</th>
                                                    </tr>
                                                <thead>
                                                <tbody >
                                                    <%
                                                    totalMoney = 0  
                                                     do While not rsDetails.eof
                                                    %>
                                                    <tr>
                                                        <td>
                                                            <img src="<%=Trim(rsDetails("imglinks"))%>" alt=""  style="aspect-ratio: 4/4;width: 50px; margin-inline: 15px;">
                                                        </td>
                                                        <td style="vertical-align: middle;"><% = Trim(rsDetails("title"))%> </td>
                                                        <td style="vertical-align: middle;"><% = Dis_str_money(Trim(rsDetails("Price")))%>?? </td>
                                                        <td style="vertical-align: middle;"> <% =rsDetails("Quantity")%></td>
                                                        <td style="vertical-align: middle;"> <% = Dis_str_money(Trim(rsDetails("UnitPrice")))%>??</td>
                                                    </tr>
                                                    <%
                                                        totalMoney = totalMoney + Trim(rsDetails("UnitPrice"))
                                                        rsDetails.movenext
                                                        Loop
                                                    %>
                                                </tbody>
                                                <tfoot>
                                                    <td colspan="5" class="w3-right-align"><span>T???ng ti???n (ch??a bao g???m ph?? ship) :</span><b><h5 class="d-inline w3-text-deep-orange"><% = Dis_str_money(totalMoney)%>??</h5></b></td>
                                                </tfoot>
                                            </table>
                                        </div>
                                        <%
                                        End if
                                        set rsDetails = nothing
                                        %>
                                        <!-- Modal footer -->
                                        <div class="modal-footer ">
                                                <button type="button" class="btn btn-danger" data-dismiss="modal">????ng</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </td>
                    </tr>
                    <% 
                        stt=stt+1
                        rs.movenext
                        Loop %>
                        
                        </table>
                </div>
                <%else %>
                <div class="w3-panel w3-pale-yellow w3-border">
                    <h3>Th??ng b??o!</h3>
                    <p>Ch??a c?? d??? li???u tr??n h??? th???ng.</p>
                </div>
            <%
                        rs.close
                        set rs=nothing
                        ENd if          
                    %>
            </div>
        </div>
        </div>
    </div>
</body>
</html>