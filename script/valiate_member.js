



function vakidate_profile_ctv() {
    $('#FChangeInf').validate({
        rules: {
            F_fullname: {
                user_name: true,
                name_minlength: 4,
                name_maxlength: 50

            },
            F_ngaysinh: {
                date_Nsinh: true,
                //date_Nsinh_fomat: true,
            },

            F_tel: {
                user_mobile: true,
                number_phoneFomat: true,
                phone_minlength: 8,
                phone_maxlength: 15

            },
            F_Address: {
                user_address: true,
            },

            F_hocvan: {
                user_hocvan: true
            },
            F_nguyenQuan: {
                user_nguyenquan: true
            },

            F_idcode: {
                user_idcode: true
            },
            F_cmtnd: {
                user_cmt: true,
                number: true,
                cmt_minlength: 9,
                cmt_maxlength: 12
            },
            F_atmbank: {
                required: true
            },


            F_nganhang: {
                user_so_tkhoan_nn: true
            },

        },

        highlight: function (element) {
            $(element).closest('.row-mg').addClass('has-error');
        },
        unhighlight: function (element) {
            $(element).closest('.row-mg').removeClass('has-error');
        },
        errorElement: 'span',
        errorClass: 'lbl-error',
        errorPlacement: function (error, element) {
            if (element.attr("name") == "F_fullname") {
                error.appendTo('#Err_name');
            }
            if (element.attr("name") == "F_ngaysinh") {
                error.appendTo('#Err_Nsinh');
            }
            if (element.attr("name") == "F_tel") {
                error.appendTo('#Err_tel');
            }
            if (element.attr("name") == "F_Address") {
                error.appendTo('#Err_Address');
            }
            if (element.attr("name") == "F_hocvan") {
                error.appendTo('#Err_hocvan');
            }
            if (element.attr("name") == "F_nguyenQuan") {
                error.appendTo('#Err_nguyenquan');
            }
            if (element.attr("name") == "F_idcode") {
                error.appendTo('#Err_idcode');
            }
            if (element.attr("name") == "F_cmtnd") {
                error.appendTo('#Err_cmtnd');
            }
            if (element.attr("name") == "F_atmbank") {
                error.appendTo('#Err_atmbank');
            }
            if (element.attr("name") == "F_nganhang") {
                error.appendTo('#Err_nganhang');
            }
        },
        submitHandler: function (form) {
            var img_name = $("div.file-footer-caption").html();
            if (img_name != "") {
                $('#F_Img').val(img_name);
            }
            $(".kv-file-upload").click();
            AdmMember_updateProfile('updateprofile');
        }

    });
}

//submitHandler: function (form) {
//    submitForm_dkmember_ctv('add_member_ctv');
//}


