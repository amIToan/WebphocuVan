<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<!--#include virtual="/include/Constant.asp"-->
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>     
        <link href="/images/logo/icon.ico" rel="icon" type="image/x-icon" />
        <link href="/images/logo/icon.ico" rel="shortcut icon" />
        <link href="/stylesheets/w3style.css" rel="stylesheet" />
        <link href="/interfaces/liberary/bootstrap4/css/bootstrap.min.css" rel="stylesheet" />
        <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/jquery-3.5.1.min.js"></script>
        <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/bootstrap.min.js"></script>
        <link rel="stylesheet" href="/selectlib/select2.min.css">
        <script src="/selectlib/select2.min.js"></script>
        <link rel="stylesheet" href="/css/sweetalert.css">
        <script src="/js/sweetalert.min.js"></script>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <title>Payment Cards</title>
    </head>
    <style>
        .select2-container .select2-selection--single{
            height: calc(1.5em + .75rem + 2px);
        }
        .hide {
            display: none !important;
        }
        .error-toan {
            color: red;
        }
    </style>
<body>
    <div class="container">
        <div class="row">
            <div class="col-lg-6 mt-3">
                <h3> Tôi là thảo mộc</h3>
                <hr>
                <p> Giỏ hàng > Thông tin giao hàng > Phương thức thanh toán</p>
                <p>Bạn đã có tài khoản? Đăng nhập</p>
                <hr>
                <form method="post" id="testJson">
                    <div class="row justify-content-center">
                      <div class="col-12  form-group">
                        <input type="text" id="fullName" name="fullName" class="form-control mb-3" placeholder="First name">
                        <div class="form-message"></div>
                      </div>
                      <div class="col-12 form-group">
                        <input type="text" id= "emailUSer" name= "emailUSer" class="form-control mb-3" placeholder="Email...">
                        <div class="form-message"></div>
                      </div>
                      <div class="col-12 form-group">
                        <input type="text" id="phoneNumber" name="phoneNumber" class="form-control mb-3" placeholder="Phone Number...">
                        <div class="form-message"></div>
                      </div>
                      <div class="col-12 mb-3">
                          <div class="row">
                                <div class="col-6 form-group">
                                    <% Call Province("Province",ProvinceID,"chosen-select  form-control","") %>
                                    <div class="form-message"></div>
                                </div>
                                <div class="col-6 form-group">
                                    <%
                                        Call District("District",0,DistrictID,"chosen-select  form-control","")  
                                        %> 
                                    <div class="form-message"></div>
                                </div>
                          </div>
                      </div>
                      <div class="col-6 form-group mb-4">
                         <% Call Wards("Ward",0,WardID,"chosen-select  form-control","") %>
                         <div class="form-message"></div>
                      </div>
                      <div class="col-6 mb-4 form-group">
                        <input type="text" class="form-control " name="Address" id="Adress" placeholder="Địa chỉ chi tiết..">
                        <div class="form-message"></div>
                      </div>
                    </div>
                    <input type="hidden" name="strProduct" id="strProduct" >
                    <div class="d-flex justify-content-between py">
                            <button type ="button" class="btn btn-light " onclick="window.location.href='/cartlists'">Giỏ hàng</button>
                            <button type="button" class="btn btn-info" data-toggle="modal" data-target="#myModal">
                                Tiến hành thanh toán
                            </button>
                            <!-- The Modal -->
                            <div class="modal fade" id="myModal">
                                <div class="modal-dialog modal-lg">
                                    <div class="modal-content">

                                        <!-- Modal Header -->
                                        <div class="modal-header">
                                            <h4 class="modal-title">Cổng thanh toán</h4>
                                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                                        </div>

                                        <!-- Modal body -->
                                        <div class="modal-body">
                                            <h5>Phương thức vận chuyển</h5>
                                            <div class="d-flex align-items-center border border-info rounded p-3">
                                                <input type="radio" name="express" id="express" style="height: 20px ;width: 20px" checked>
                                                <label for="express" class="m-0 p-2"> Giao hàng tận nơi</label>
                                            </div>
                                            <div class="p-2"></div>
                                            <h5>Phương thức thanh toán</h5>
                                            <div class="border border-info rounded p-3" id="payingMethods">
                                                <div class=" d-flex align-items-center  ">
                                                    <input type="radio" name="postMethod" id="postMethod" value="0" style="height: 20px ;width: 20px" checked>
                                                    <label for="postMethod" class="m-0 p-2">
                                                        <img src="/images_upload/IMG_Customer/cod.svg" alt="">
                                                         Thanh toán khi giao hàng
                                                    </label>
                                                </div>
                                                <div class=" d-flex text-center p-3 border-top border-info tab-content "
                                                    style="margin-inline: -1rem; background-color: #fafafa;">
                                                    <span>
                                                        Cảm ơn bạn đã tin dùng mua hàng tại Tôi Là Thảo Mộc.
                                                        Chúng tôi sẽ sớm liên hệ với bạn để Xác Nhận Đơn Hàng qua điện thoại trước khi giao
                                                        hàng!
                                                    </span>
                                                </div>
                                                <hr class="border-top border-info " style="margin-inline: -1rem; margin-block: 0px">
                                                <div class="d-flex align-items-center ">
                                                    <input type="radio" name="interbanking" id="interbanking" value="1" style="height: 20px ;width: 20px">
                                                    <label for="interbanking" class="m-0 p-2">
                                                        <img src="/images_upload/IMG_Customer/other.svg" alt="">
                                                        Thanh toán bằng phương thức Banking
                                                    </label>
                                                </div>
                                                <div class="text-center p-3 border-top border-info tab-content  hide"
                                                    style="margin-inline: -1rem;margin-bottom: -1rem; background-color: #fafafa;">
                                                    <span class="d-block">
                                                        <% if Bank <>"" then 
                                                            Response.Write("* Ngân hàng :" & Bank)
                                                            else
                                                            %>
                                                        *Ngân hàng MB BAnk - Chi Nhánh Mỹ Đình
                                                        <%End If%>
                                                    </span>
                                                    <span class="d-block">
                                                        <% if STK <>"" then 
                                                            Response.Write("* Tên Chủ TK: " & Tentaikhoan)
                                                            else
                                                            %>
                                                        *Tên Chủ TK: Tạ Quang Toản
                                                        <%End If%>
                                                    </span>
                                                    <span class="d-block">
                                                        <% if STK <>"" then 
                                                            Response.Write("* Số tài khoản: " & STK )
                                                            else
                                                            %>
                                                        *Số TK: 888841098888
                                                        <%End If%>
                                                    </span>
                                                    <span class="d-block">
                                                        <% if Tel_sys <>"" then 
                                                            Response.Write("* Số Điện Thoại: " & Tel_sys)
                                                            else
                                                            %>
                                                        *Điện Thoại: 0389952255
                                                        <%End If%>
                                                    </span>
                                                    <span class="d-block">
                                                        -------------------------------------------------------------------------
                                                    </span>
                                                    <span class="d-block">
                                                        Nội dung chuyển khoản vui lòng ghi rõ:
                                                    </span>
                                                    <span class="d-block">
                                                        Họ Tên người chuyển - Sản phẩm đã mua - Số Điện Thoại
                                                    </span>
                                                </div>
                                            </div>
                                        </div>
                                        <!-- Modal footer -->
                                        <div class="modal-footer justify-content-between">
                                                <button type="button" class="btn btn-light" data-dismiss="modal">Quay lại</button>
                                                <button type="submit" class="btn btn-info" form="testJson">Hoàn tất đơn hàng</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                    </div>
                  </form>
            </div>
            <div class="col-lg-6 mt-3" style="background-color: #fafafa">
                <h3> Thông tin đơn hàng</h3>
                <hr>
                <h5 id="Payment-button" style="background-color: #3f7e3b; color:white; display: inline-block; padding: 10px; border-radius: 5px; cursor: pointer"> <i class="fas fa-cart-plus me-1"></i> Ẩn thông tin đơn hàng</h5>
                <div id="Payment-control">
                    <ul style="padding:0 ; list-style: none; margin: 0;" id="productContainer">
                    
                    </ul>
                    <div class="row">
                        <p class="col-9 text-danger"> Khuyến mại : </p>
                        <p id="temAmount" class="col-3 text-danger" ></p>
                    </div>
                    <p> Phí vận chuyển : (tùy địa điểm)</p>
                    <hr>
                    <div class="d-flex justify-content-between" >
                        <h5 class="col-9">Tổng tiền : </h5>
                        <h5 class="col-3" id="totalMoney"></h5>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <script src="/javascript/valid_control.js"></script>
    <script>
        let productsInCart = JSON.parse(localStorage.getItem('shoppingCart'));
        document.getElementById("strProduct").value= JSON.stringify(productsInCart)
        const productContainer = document.getElementById("productContainer")
        let products = productsInCart.map( product  => {
            let Discount ; 
            if (product.discountValue) {
                Discount = product.discountValue;
            }
            return `
				<li class="row align-items-center justify-content-between">
					<div class="col-9 row align-items-center">
                            <img src="${product.image}" alt="${product.name}"  style="aspect-ratio: 4/4;width: 100px; margin-inline: 15px;">
                            <div>${product.name} 
                                <p> Số lượng: ${product.count}</p>
                                <div class="text-danger"><i>${Discount ? "Khuyến mại : " +  Discount + " %" : "Không có khuyến mại" }</i></div>
                            </div>
                        </div>
                        <div class="col-3 row align-items-center priceValue" >${product.price.toLocaleString() + "đ"}</div>
				</li>
                <hr>
                ` 
        })
        productContainer.innerHTML = products.join("");
        const result =  productsInCart.reduce( (acc, item, index) => {
            if (item.discountValue) {
                return acc += (item.price - (item.price /100) * item.discountValue ) 
            }else {
                return acc += item.price
            }
        },0)
        const discountedResult =  productsInCart.reduce( (acc, item, index) => {
            if (item.discountValue) {
                return acc += ((item.price /100) * item.discountValue ) 
            }else {
                return acc 
            }
        },0)
        const totalMoney = document.getElementById("totalMoney");
        const temMoney = document.getElementById("temAmount");
        totalMoney.innerText = result.toLocaleString() + "đ";
        temMoney.innerText = "- " + discountedResult.toLocaleString() + "đ";
        $("#Payment-button").click(function(){
            $("#Payment-control").toggle("slow");
        });
        function get_District(_ProvinceID) {
            $.ajax({
                method: "get",
                url: "/include/test_province.asp",
                data: {
                    keyValue : "province", 
                    ProvinceID: _ProvinceID
                },
                contentType: "application/json",
                success: function (obj) {
                    var t = $.parseJSON(obj);
                    if (t.error.status) {
                        $("#District").append(t.data);
                    }
                },
                error: function (t) {
                    //console.log("getCmsProductHot=>error", JSON.stringify(t))
                }
            })
        }

        function get_Ward(_DistrictID) {
            $.ajax({
                method: "get",
                url: "/include/test_province.asp",
                data: {
                    keyValue : "district",
                    DistrictID: _DistrictID
                },
                contentType: "application/json",
                success: function (obj) {
                    var t = $.parseJSON(obj);
                    if (t.error.status) {
                        $("#Ward").append(t.data);
                    }
                },
                error: function (t) {
                    //console.log("getCmsProductHot=>error", JSON.stringify(t))
                }
            })
        }
         $(document).ready(function() {
                    // change select option
                $('#Province').select2();
                $('#District').select2();
                $('#Ward').select2();
                // get Province
                $("#Province").change(function () {
                    $('#District').find('option').remove().end()
                    $('#District').append('<option value="0">Chọn quận / huyện</option>');
                    get_District($("select[name=Province] option:selected").val())
                    $("select[name=Province] option:selected").val()
                });

                $("#District").change(function () {
                    $('#Ward').find('option').remove().end()
                    $('#Ward').append('<option value="0">Chọn phường xã</option>');
                    get_Ward($("select[name=District] option:selected").val())
                });
                // $('form').on('submit', function (event) {
                //     event.preventDefault();
                //     event.stopPropagation();
                    
                // })
            })
            const payingMethods = document.querySelector("#payingMethods")
            // document.querySelector("#express").checked = true;
            payingMethods.addEventListener("click", (e)=> {
                const radioButton = e.target.closest("input[type='radio']");
                if (!radioButton) return;
                console.log(radioButton);
                const allradioBtn = Array.from(payingMethods.querySelectorAll("input[type='radio']")) ;
                const tabContent = Array.from(payingMethods.querySelectorAll(".tab-content"));
                tabContent.forEach( item => {
                        item.classList.add("hide");
                })
                allradioBtn.forEach( (item,index) => {
                    item.checked = false;
                    radioButton.checked = true;
                    tabContent[allradioBtn.indexOf(radioButton)].classList.remove("hide");
                })
                ;
            })
            Validator({
                form: '#testJson',
                formGroupSelector: '.form-group',
                errorSelector: '.form-message',
                styleError: 'error-toan',
                rules: [
                    Validator.isfullText('#fullName'),
                    Validator.isEmail('#emailUSer'),
                    Validator.isphoneNumber('#phoneNumber'),
                    Validator.isphoneNumber('#phoneNumber'),
                    Validator.isRequired('#Province'),
                    Validator.isRequired('#District'),
                    // Validator.isRequired('#Ward'),
                    Validator.isRequired('#Adress')
                ],
                onSubmit: function (data) {
                        $.ajax({
                            url: '/include/ajax_payingcart.asp',
                            type: 'POST',
                            data: $("#testJson").serialize() + "&keyword=" + "payingProduct",
                            async: false,
                            cache: false,
                            success: function (returndata) {
                                if (returndata.status == 1) {
                                    swal({
                                        title: "Chúc mừng bạn đã đặt hàng thành công",
                                        text: `Mã đơn hàng của bạn là ${returndata.order} . Chúng tôi sẽ gửi hàng nhanh nhất có thể cho bạn.`,
                                        text: "Vui lòn gọi Hotline : <%=Hotline%> để được hỗ trợ. Xin cảm ơn",
                                        type: "success",
                                        showCancelButton: false,
                                        confirmButtonColor: "#DD6B55",
                                        confirmButtonText: "OK!",
                                        closeOnConfirm: false,
                                        showLoaderOnConfirm: true
                                    }, function () {
                                        localStorage.removeItem("shoppingCart");
                                        window.location.href="/index.asp";
                                    });
                                } else {
                                    alert("Đã có lỗi xảy ra. Vui lòng nhập lại!!!")
                                    window.location.reload();
                                }
                            }
                        });
                }
            })
    </script>
</body>
</html>
<% 
Sub Province(ProvinceName,ProvinceID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select * From Province Order by Orderby" 
    'Response.write sql  
	rs.Open sql, con, 1
	
    response.Write "<select data-placeholder='Chọn Tỉnh/ Thành phố'  class='"&class_&"' style='"&style_&"' name="&ProvinceName&" id="&ProvinceName&">"
	response.Write"<option value='0'>Chọn Tỉnh/ Thành phố</option>"
    Do while not rs.eof
		response.Write"<option value=""" & clng(rs("ProvinceID"))  & """"
		if clng(rs("ProvinceID"))=ProvinceID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("NameProvince")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>
<% 
Sub District(DistrictName,ProvinceID,DistrictID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select DistrictID,NameDistrict From Province_district Where ProvinceID="&ProvinceID
    Response.write ("<p style='display:none'>"&sql&"</p>")
    'Response.write sql
	rs.Open sql, con, 1

	response.Write "<select data-placeholder='Chọn Quận/ Huyện'  class='"&class_&"' style='"&style_&"' name="&DistrictName&" id="&DistrictName&">"
	Response.write "<option value='0'>Chọn quận / huyện</option>"
    Do while not rs.eof
		response.Write "<option value=""" & clng(rs("DistrictID"))  & """"
		if clng(rs("DistrictID"))=DistrictID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("NameDistrict")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>

<% 
Sub Wards(WardName,DistrictID,WardID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select wardID,wardName From Province_ward Where DistrictID="&DistrictID
    'Response.write sql
	rs.Open sql, con, 1
	response.Write "<select data-placeholder='Chọn phường xã'  class='"&class_&"' style='"&style_&"' name="&WardName&" id="&WardName&">"
	Response.write "<option value='0'>Chọn phường xã</option>"
    Do while not rs.eof
		response.Write "<option value=""" & clng(rs("wardID"))  & """"
		if clng(rs("wardID"))=WardID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("wardName")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>


