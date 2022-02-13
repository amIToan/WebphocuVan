function closeCart() {
	const cart = document.querySelector('.producstOnCart');
	cart.classList.toggle('hide');
	document.querySelector('body').classList.remove('stopScrolling')
}
const openShopCart = document.querySelector('.shoppingCartButton');
openShopCart.addEventListener('click', () => {
	const cart = document.querySelector('.producstOnCart');
	cart.classList.toggle('hide');
	document.querySelector('body').classList.toggle('stopScrolling');
});
const closeShopCart = document.querySelector('#closeButton');
const overlay = document.querySelector('.overlay');
closeShopCart.addEventListener('click', closeCart);
overlay.addEventListener('click', closeCart);
let productsInCart = JSON.parse(localStorage.getItem('shoppingCart')); // hàm để tạo array 
if (!productsInCart) {
	productsInCart = [];
}
const parentElement = document.querySelector('#buyItems'); // thanh UI để lưu lại sẩn phẩm
const cartSumPrice = document.querySelector('#sum-prices'); // lấy ra cái giá cảu nó
const products = document.querySelectorAll('.product-under'); // chọn ra tất cả khung bo products
const countTheSumPrice = function () { // 4
	let sum = 0;
	productsInCart.forEach(item => {
		sum += item.price;
	});
	return sum;
}
const updateShoppingCartHTML = function () {  // 3
	localStorage.setItem('shoppingCart', JSON.stringify(productsInCart)); // tạo ra localstorage với dữ liệu là array ProductsInCart 
	if (productsInCart.length > 0) { // Nếu có dữ liệu thì nó sẽ nhảy vào đây
		let result = productsInCart.map(product => { // tóm lại là lặp qua array lưu và result
			return `
				<li class="buyItem">
					<img src="${product.image}">
					<div>
						<h4>${product.name}</h4>
						<h6>${product.price.toLocaleString() + "đ"}</h6>
						<div>
							<button class="button-minus" data-id=${product.id}>-</button>
							<span class="countOfProduct">${product.count}</span>
							<button class="button-plus" data-id=${product.id}>+</button>
						</div>
					</div>
				</li>`
		});
		parentElement.innerHTML = result.join(''); // parentEle chính là ul và result.join sẽ ra đc các li bên trong 
		document.querySelector('.checkout').classList.remove('hidden');
		cartSumPrice.innerHTML = countTheSumPrice().toLocaleString() + "đ"; // chỗ này sẽ lưu tổng giá vào rổ
	}
	else { // Nếu mà array dữ liệu không có gì thì là như thế này 
		document.querySelector('.checkout').classList.add('hidden');
		parentElement.innerHTML = '<h4 class="empty">Your shopping cart is empty</h4>';
		cartSumPrice.innerHTML = '';
	}
}

function updateProductsInCart(product) { // 2
	for (let i = 0; i < productsInCart.length; i++) { // lặp qua vòng array 
		if (productsInCart[i].id == product.id) { // nếu id của product bằng với aray của localstorage thì 
			productsInCart[i].count += product.count; // cái số lượng của nó sẽ thêm 1 
			productsInCart[i].price = productsInCart[i].basePrice * productsInCart[i].count; // cái giá cảu nó sẽ bằng cái giá base X lên counts
			return;
		}
	}
	productsInCart.push(product); // Nếu k có thì nó sẽ push product 
}

products.forEach(item => {   // 1
	item.addEventListener('click', (e) => {
		if (e.target.classList.contains('addToCart')) {
			const productID = e.target.dataset.productId;
			document.querySelector('.producstOnCart').classList.remove('hide');
			const productName = item.querySelector('.productName').innerHTML;
			const productPrice = parseInt(item.querySelector('.priceValue').innerHTML.replace('.', ''));
			const productImage = item.querySelector('img').src;
			let productCount = item.querySelector('input[type="number"]')
			let productPrice2;
			if (productCount && productCount.value) {
				productCount = parseInt(productCount.value)
				productPrice2 = productCount * productPrice;
			} else {
				productCount = 1
			}
			let product = {
				name: productName,
				image: productImage,
				id: productID,
				count: +productCount,
				price: productPrice2 ? productPrice2 : productPrice,
				basePrice: +productPrice,
			}
			console.log(product.count)
			updateProductsInCart(product); // gọi hàm để thực hiện upgrade- túm lại là đẩy product vào hoặc nếu có và + giá
			updateShoppingCartHTML();
		}
	});
});

parentElement.addEventListener('click', (e) => { // Last
	const isPlusButton = e.target.classList.contains('button-plus');
	const isMinusButton = e.target.classList.contains('button-minus');
	if (isPlusButton || isMinusButton) {
		for (let i = 0; i < productsInCart.length; i++) {
			if (productsInCart[i].id == e.target.dataset.id) {
				if (isPlusButton) {
					productsInCart[i].count += 1
				}
				else if (isMinusButton) {
					productsInCart[i].count -= 1
				}
				productsInCart[i].price = productsInCart[i].basePrice * productsInCart[i].count;

			}
			if (productsInCart[i].count <= 0) {
				productsInCart.splice(i, 1);
			}
		}
		updateShoppingCartHTML();
	}
});

updateShoppingCartHTML();
