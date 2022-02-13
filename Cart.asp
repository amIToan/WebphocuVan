
<html>
<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/Constant.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<head>
    <title>Giỏ hàng của bạn</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />     
    <link href="/images/logo/icon.ico" rel="icon" type="image/x-icon" />
    <link href="/images/logo/icon.ico" rel="shortcut icon" />
    <link href="/interfaces/liberary/bootstrap4/css/bootstrap.min.css" rel="stylesheet" />
    <script type="text/javascript" src="/interfaces/liberary/bootstrap4/js/jquery-3.5.1.min.js"></script>
    <link rel="stylesheet" href="/css/sweetalert.css">
    <script src="/js/sweetalert.min.js"></script>
    <link href="/stylesheets/w3style.css" rel="stylesheet" />
    <script type="text/javascript" src="/interfaces/owlcarousel/owl.carousel.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <link href="/interfaces/fonts/font-awesome.css" rel="stylesheet" />
    <link href="/interfaces/css/ceo_lam.css" rel="stylesheet" />
    <link href="/interfaces/css/bootstrap.css" rel="stylesheet" />
    <%Call code_google() %>
</head>
<body>
<!--#include virtual="/include/func_common.asp"-->
<!--#include virtual="/include/Fs_cotruct.asp"-->
<!--#include virtual="/include/func_tiny.asp"-->
<!--#inclue virtual="/include/function_toan.asp" -->
    <%
    Call Header() 
    Call Fs_menuMOblie()
    Call Fs_menu()
    %>
    <div class="container py-5">
        <h4>Giỏ hàng của bạn </h4>
        <div class="w3-responsive">
            <table class="w3-table-all w3-centered w3-hoverable" >
                <thead >
                    <tr style="background: #3f7e3b">
                        <th style="min-width: 100px">Sản phẩm </th>
                        <th style="min-width: 350px"> Mô tả</th>
                        <th style="min-width: 100px"> Giá </th>
                        <th style="min-width: 150px"> Số lượng</th>
                        <th>Tổng</th>
                        <th>Xóa</th>
                    </tr>
                <thead>
                <tbody id="tableproContainer">
                    
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="3" class="w3-left-align">
                        <form method="post" id="cartlists">
                           <div class="d-flex justify-content-between">
                               <input type="text" name="discountedKey" class="form-control" style="width: 75%"/> 
                                <button type ="submit" class="btn btn-info">Nhập mã giảm giá</button>
                           </div> 
                        </td>
                        </form>
                        <td colspan="3" class="w3-right-align"><span>Tổng tiền (chưa bao gồm phí ship) :</span><b><h4 class="d-inline w3-text-deep-orange" id="totalMoney"></h4></b></td>
                    </tr>
                </tfoot>
            </table>
        </div>
        <div class="d-flex justify-content-between py-4">
            <button class="btn btn-primary" onclick="window.location.href='/index.asp'">Tiếp tục mua hàng</button>
            <button class="btn btn-primary" onclick="window.location.href='/payingcart'">Đặt hàng</button>
        </div>
    </div>
    <%
    call Orderdetails()
    Call  Fs_Footer()
    call backTop()
    Call Item_support_Toan(lang)
    %>
    <script src="/javascript/shopping-cart.js"></script>
    <script>
        const productContainer = document.getElementById("tableproContainer")
        const updatetableCartHTML = function () {
            localStorage.setItem('shoppingCart', JSON.stringify(productsInCart));
            if (productsInCart.length > 0) { // Nếu có dữ liệu thì nó sẽ nhảy vào đây
             let proTable = productsInCart.map( (product, index)  => {
            return `
				<tr>
                    <td style="vertical-align: middle;">
                    <img src="${product.image}" alt="${product.name}"  style="aspect-ratio: 4/4;width: 50px; margin-inline: 15px;">
                    </td>
                    <td style="vertical-align: middle;">${product.name}</td>
                    <td style="vertical-align: middle;">${(product.price / product.count).toLocaleString() + "đ"}</td>
                    <td style="vertical-align: middle;">
                        <button class="btn button-minus" data-id=${product.id}>-</button>
						<span class="countOfProduct">${product.count}</span>
						<button class="btn button-plus" data-id=${product.id}>+</button>
                    </td >
                    <td style="vertical-align: middle;">${product.price.toLocaleString() + "đ"}</td>
                    <td style="vertical-align: middle;">
                        <button class="btn btn-danger deleteBtn" data-id=${product.id} data-index=${index}>Xóa</button>
                    </td>
                </tr>
                ` 
            })
            productContainer.innerHTML = proTable.join(''); // parentEle chính là ul và result.join sẽ ra đc các li bên trong 
            cartSumPrice.innerHTML =  countTheSumPrice().toLocaleString() + "đ"; // chỗ này sẽ lưu tổng giá vào rổ
            }
            else { // Nếu mà array dữ liệu không có gì thì là như thế này 
                productContainer.innerHTML = 'lol';
                cartSumPrice.innerHTML = '';
            }
        }
        updatetableCartHTML();
        productContainer.addEventListener('click', (e) => { // Last
            const isPlusButton = e.target.classList.contains('button-plus');
            const isMinusButton = e.target.classList.contains('button-minus');
            const deleteButton = e.target.closest(".deleteBtn");
            if (!(isPlusButton || isMinusButton || deleteButton)) return;
            if (isPlusButton || isMinusButton || deleteButton ) {
                for (let i = 0; i < productsInCart.length; i++) {
                    if (productsInCart[i].id == e.target.dataset.id) {
                        if (isPlusButton) {
                            productsInCart[i].count += 1
                            productsInCart[i].price = productsInCart[i].basePrice * productsInCart[i].count;
                        }
                        else if (isMinusButton) {
                            productsInCart[i].count -= 1
                            productsInCart[i].price = productsInCart[i].basePrice * productsInCart[i].count;
                            if (productsInCart[i].count <= 0) {
                            productsInCart[i].count = 1;
                            productsInCart[i].price = productsInCart[i].basePrice * productsInCart[i].count;
                            }
                        } else if (deleteButton) {
                            productsInCart.splice(e.target.dataset.index, 1)
                            console.log(productsInCart)
                        }
                    }
                  
                }
                updatetableCartHTML();
                updateShoppingCartHTML();
                getTotalMoney();
            }
        });
        function getTotalMoney() {
            result =  productsInCart.reduce( (acc, item, index) => {
            return acc += item.price
            },0)
            totalMoney.innerText = result.toLocaleString() + "đ";
        }
        let totalMoney = document.getElementById("totalMoney");
        let result =  productsInCart.reduce( (acc, item, index) => {
            return acc += item.price
            },0)
        totalMoney.innerText = result.toLocaleString() + "đ";
        $('#cartlists').on('submit', function (event) {
                event.preventDefault();
                event.stopPropagation();
                        $.ajax({
                            url: '/include/ajax_payingcart.asp',
                            type: 'POST',
                            data: $("#cartlists").serialize() + "&keyword=" + "cartlists",
                            async: false,
                            cache: false,
                            success: function (returndata) {
                                if (returndata.status == 1) {
                                    swal({
                                        title: "Chúc mừng bạn sử dụng mã thành công",
                                        text: "Vui lòng thanh toán hóa đơn để nhận hàng nhanh nhất !",
                                        type: "success",
                                        showCancelButton: false,
                                        confirmButtonColor: "#DD6B55",
                                        confirmButtonText: "OK!",
                                        closeOnConfirm: false,
                                        showLoaderOnConfirm: true
                                    }, function () {
                                        const discountedValue = parseInt(returndata.discountVal)
                                        productsInCart.forEach( item => {
                                            item.discountValue = discountedValue
                                        })
                                        localStorage.setItem('shoppingCart', JSON.stringify(productsInCart)); 
                                        window.location.href='/payingcart';
                                    });
                                } else {
                                    alert("Đã có lỗi xảy ra. Vui lòng nhập lại!!!")
                                    console.log(productsInCart)
                                    window.location.reload();
                                }
                            }
                        });        
        })
    </script>
</body>
</html>