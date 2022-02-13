

//--------------------------------------------------------------------------------------------------------------
function Res_setup(_key, _id) {
    //window.open("/outs/" + _alias + "/" + _key + "/" + _id + ".ios") // mo ra trang moi
    window.location.assign("/out/" + _key + "/" + _id + ".ios") // mo ra trang moi
}


function Func_resgister(_key, _id) {
    $.ajax({
        url: "/mod/" + _key + "/" + _id+"", 
        data: $("#FRegister").serialize() + "&_key=" + _key + "&_id=" + _id,
        type: "POST",
        cache: false,
        dataType: "html",
        success: function (rs) {
            if (rs == "1") {
                swal({
                    title: "Hệ Thống",
                    text: "Gửi thành công. Xin cảm ơn.",
                    type: "success",
                    timer: 3000,
                    closeOnConfirm: false
                });

            }
            else {
                swal({
                    title: "Hệ thống",
                    text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                    type: "error",
                    timer: 3000,
                    showConfirmButton: false
                });
            }
        },

        error: function (rs) {
            swal({
                title: "Hệ thống",
                text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                type: "error",
                timer: 3000,
                showConfirmButton: false
            });
        },


    });

    setTimeout(function () {
        window.location.reload(1);
    }, 3000);
}

function Func_resEmail(_key, _id) {
    $.ajax({
        url: "/mod/" + _key + "/" + _id + "",
        data: $("#Fremail").serialize() + "&_key=" + _key + "&_id=" + _id,
        type: "POST",
        cache: false,
        dataType: "html",
        success: function (rs) {
            if (rs == "1") {
                swal({
                    title: "Hệ Thống",
                    text: "Gửi thành công. Xin cảm ơn.",
                    type: "success",
                    timer: 3000,
                    closeOnConfirm: false
                });

            }
            else {
                alert(rs);
                swal({
                    title: "Hệ thống",
                    text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                    type: "error",
                    timer: 3000,
                    showConfirmButton: false
                });
            }
        },

        error: function (rs) {

            alert(rs);
            swal({
                title: "Hệ thống",
                text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                type: "error",
                timer: 3000,
                showConfirmButton: false
            });
        },


    });

    setTimeout(function () {
        window.location.reload(1);
    }, 3000);
}



function Func_Filldata(_key, _id) {
    $.ajax({
        url: "/mod/" + _key + "/" + _id + "",
        data:"_key=" + _key + "&_id=" + _id,
        type: "POST",
        cache: false,
        dataType: "html",
        success: function (rs) {


            $("#FillData").html(rs);
        },

        error: function (rs) {
            swal({
                title: "Hệ thống",
                text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                type: "error",
                timer: 3000,
                showConfirmButton: false
            });
        },


    });


}

function Func_FillProduct(_key, _id) {
    $.ajax({
        url: "/mod/" + _key + "/" + _id + "",
        data: "_key=" + _key + "&_id=" + _id,
        type: "POST",
        cache: false,
        dataType: "html",
        success: function (rs) {


            $("#FillProduct").html(rs);
        },

        error: function (rs) {
            swal({
                title: "Hệ thống",
                text: "Xin lỗi quý khách, hệ thống đang bảo trì.",
                type: "error",
                timer: 3000,
                showConfirmButton: false
            });
        },


    });


}









function Fomat_Mony(num) {
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1,")
}


