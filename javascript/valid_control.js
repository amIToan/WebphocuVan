// Đối tượng `Validator`
function Validator(options) {
    function getParent(element, selector) {
        while (element.parentElement) {
            if (element.parentElement.matches(selector)) {
                return element.parentElement;
            }
            element = element.parentElement;
        }
    }
    var selectorRules = {};

    // Hàm thực hiện validate
    function validate(inputElement, rule) {
        var errorElement = getParent(inputElement, options.formGroupSelector).querySelector(options.errorSelector);
        var errorMessage;

        // Lấy ra các rules của selector
        var rules = selectorRules[rule.selector];

        // Lặp qua từng rule & kiểm tra
        // Nếu có lỗi thì dừng việc kiểm
        for (var i = 0; i < rules.length; ++i) {
            switch (inputElement.type) {
                case 'radio':
                case 'checkbox':
                    errorMessage = rules[i](
                        formElement.querySelector(rule.selector + ':checked')
                    );
                    break;
                default:
                    errorMessage = rules[i](inputElement.value);
            }
            if (errorMessage) break;
        }

        if (errorMessage) {
            errorElement.innerText = errorMessage;
            getParent(inputElement, options.formGroupSelector).classList.add(options.styleError);
            setTimeout(function () {
                errorElement.innerText = '';
                getParent(inputElement, options.formGroupSelector).classList.remove(options.styleError);
            }, 10000)
        } else {
            errorElement.innerText = '';
            getParent(inputElement, options.formGroupSelector).classList.remove(options.styleError);
        }

        return !errorMessage;
    }

    // Lấy element của form cần validate
    var formElement = document.querySelector(options.form);
    if (formElement) {
        //////Khi submit form
        formElement.onsubmit = function (e) {
            e.preventDefault();
            var isFormValid = true;
            // Lặp qua từng rules và validate
            options.rules.forEach(function (rule) {
                var inputElement = formElement.querySelector(rule.selector);
                //console.log(inputElement)
                if (!inputElement.disabled) {
                    var isValid = validate(inputElement, rule); // Neu ma co loi thi return ve !errormessage nghia la dang false , va bay h minh chuyen thanh cho la true de mà hiểu nôm na true nghĩa là lại có lỗi 
                    if (!isValid) {
                        isFormValid = false;
                    }
                }
            });

             if (isFormValid) {
                 // Trường hợp submit với javascript
                 if (options.onSubmit && typeof options.onSubmit === 'function') {
                     var enableInputs = formElement.querySelectorAll('[name]:not([disabled]');
                     var formValues = Array.from(enableInputs).reduce(function (values, input) {
                         switch (input.type) {
                             case 'radio':
                                 const valueOfInput = formElement.querySelector(`input[name="${input.name}"]`)
                                 if (valueOfInput.checked) {
                                    values[input.name] = valueOfInput.value;
                                 }
                                 break;
                             case 'checkbox':
                                 if (!input.matches(':checked')) {
                                     values[input.name] = '';
                                     return values;
                                 }
                                 if (!Array.isArray(values[input.name])) {
                                     values[input.name] = [];
                                 }
                                 values[input.name].push(input.value);
                                 break;
                             case 'file':
                                 values[input.name] = input.files;
                                 break;
                             default:
                                 values[input.name] = input.value;
                         }

                         return values;
                     }, {});
                     options.onSubmit(formValues);
                 } else{
                    formElement.submit()
                 }
             }
             
         }
         //////Lặp qua mỗi rule và xử lý (lắng nghe sự kiện blur, input, ...)
        options.rules.forEach(function (rule) {
            const inputElement = formElement.querySelectorAll(rule.selector);
            const arrayTest =  Array.from(inputElement)
            const arrayNotDisabile = arrayTest.filter( item => !item.disabled);
            // Lưu lại các rules cho mỗi input
            if (Array.isArray(selectorRules[rule.selector])) {
                selectorRules[rule.selector].push(rule.test);
            } else {
                selectorRules[rule.selector] = [rule.test];
            }

            var inputElements = formElement.querySelectorAll(rule.selector);

            arrayNotDisabile.forEach(function (inputElement) {
                // Xử lý trường hợp blur khỏi input
                inputElement.onblur = function () {
                    validate(inputElement, rule);
                }

                // Xử lý mỗi khi người dùng nhập vào input
                inputElement.oninput = function () {
                    var errorElement = getParent(inputElement, options.formGroupSelector).querySelector(options.errorSelector);
                    errorElement.innerText = '';
                    getParent(inputElement, options.formGroupSelector).classList.remove('error');
                }
            });
        });
    }

}
// Định nghĩa rules
// Nguyên tắc của các rules:
// 1. Khi có lỗi => Trả ra message lỗi
// 2. Khi hợp lệ => Không trả ra cái gì cả (undefined)
Validator.isRequired = function (selector, message) {
    return {
        selector: selector,
        test: function (value) {
            if (value == 0 || value == 'undefined' || value =='null') return message || 'Vui lòng nhập giá trị đúng cho trường này.';
            return value ? undefined : message || 'Vui lòng nhập giá trị đúng cho trường này.'
        }
    };
}
Validator.isRequiredAlphanumericCharacters = function (selector, message) {
    
    return {
        selector: selector,
        test: function (value) {
            var letters = /^[0-9a-zA-Z]+$/;
            return letters.test(value) ? undefined : message || 'Vui lòng nhập tên của bạn.'
        }
    };
}
//var letters = /^[A-Za-z]+$/; if(inputtxt.value.match(letters))
Validator.isfullText = function (selector, message) {
    return {
        selector: selector,
        test: function (value) {
            var letters = /^[A-Za-z]+$/;
            var lettter2 = /[,#-\/\s\!\@\$.....]/gi;
            if (value.match(letters) || value.match(lettter2)) {
                return undefined;
            }else {
                return  message || 'Vui lòng nhập tên của bạn.'
            }
        }
    };
}
Validator.isEmail = function (selector, message) {
    return {
        selector: selector,
        test: function (value) {
            var regex = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
            return regex.test(value) ? undefined : message || 'Trường này phải là email';
        }
    };
}
Validator.isNumber = function (selector, message) {
    return {
        selector: selector,
        test: function (value) {
            var numbers = /^(?!0\d)\d*(\.\d+)?$/;
            return numbers.test(value) ? undefined : message || 'Trường này phải là các chữ số';
        }
    };
}
Validator.minLength = function (selector, min, message) {
    return {
        selector: selector,
        test: function (value) {
            return value.length >= min ? undefined : message || `Vui lòng nhập tối thiểu ${min} kí tự`;
        }
    };
}
Validator.isConfirmed = function (selector, getConfirmValue, message) {
    return {
        selector: selector,
        test: function (value) {
            return value === getConfirmValue() ? undefined : message || 'Giá trị nhập vào không chính xác';
        }
    }
}
Validator.isphoneNumber = function (selector, message) {
    return {
        selector: selector,
        test: function (value) {
            //if (!value || void 0 === value) return 'dit me may';
            if (value.length > 11) return 'Bạn đã nhập quá dãy số quá dài.';
            if (!value.startsWith("0")) return message || 'Vui lòng nhập đúng mã vùng';
            var numbers = /^(09|08|03|07|05|04)(\d{8})$/;
            return numbers.test(value) ? undefined : message || 'Trường này phải là các chữ số';
        }
    }
}