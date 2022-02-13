
$(document).ready(function () {
    $("#fcompany").submit(function (e) {
        e.preventDefault();
    });
});

function Navi_base(_key, _id) {
    var formData = new FormData($("#fcompany")[0]);
    $.ajax({
        type: "POST",
        data: formData,
        url: "/administrator/system/Fs_base.asp",
        cache: false,
        async: false,
        contentType: false,
        processData: false,
        success: function (rs) {
            //console.log(rs);
            if (rs == "1") {
                swal({
                    title: "Hệ Thống",
                    text: "Tạo thành công.",
                    type: "success",
                    timer: 3000,
                    closeOnConfirm: false
                });
            }
            else {
                swal({
                    title: "Hệ thống",
                    text: "Ôi hỏng.",
                    type: "error",
                    timer: 3000,
                    showConfirmButton: false
                });
            }
        },
        error: function (rs) {
            //console.log(rs);
            //alert(rs.responseText);
            swal({
                title: "",
                text: "Lỗi rồi",
                type: "error",
                timer: 2000,
                showConfirmButton: false
            });
        },
    });
    return false;
}
function Navi_baseUpdate(_key, _id) {
    var formData = new FormData($("#F_baseEdit")[0]);
    $.ajax({
        type: "POST",
        data: formData,
        url: "/administrator/system/Fs_base.asp",
        cache: false,
        async: false,
        contentType: false,
        processData: false,
        success: function (rs) {
            //console.log(rs);
            if (rs == "1") {
                swal({
                    title: "Hệ Thống",
                    text: "Thành công.",
                    type: "success",
                    timer: 3000,
                    closeOnConfirm: false
                });
            }
            else {
                swal({
                    title: "Hệ thống",
                    text: "Ôi hỏng.",
                    type: "error",
                    timer: 3000,
                    showConfirmButton: false
                });
            }
        },
        error: function (rs) {
            //console.log(rs);
            //alert(rs.responseText);
            swal({
                title: "",
                text: "Lỗi rồi",
                type: "error",
                timer: 2000,
                showConfirmButton: false
            });
        },
    });
    return false;
}



function Navi_baseEdit(_key, _id, _idLang) {
    this.window.location.href = "/base/" + _key + "/" + _id + ".html";
}


function Navi_baseDel(_key, _id, _idLang) {
    swal({
        title: "?",
        text: "You will not be able to recover this imaginary file!",
        type: "warning",
        showCancelButton: true,
        confirmButtonColor: "#F00",
        confirmButtonText: "Yes!",
        cancelButtonText: "No",
        closeOnConfirm: false,
        closeOnCancel: false
    }, function (isConfirm) {
        if (isConfirm) {
            
            $.ajax({
                type: "POST",
                url: "/administrator/system/Fs_baseDel.asp",
                data: {
                    "_key": _key,
                    "lamgG": _idLang,
                    "_id": _id
                },
                cache: false,
                dataType: "html",
                success: function (rs) {
                    if (rs == "1") {
                        swal({
                            title: "Hệ Thống",
                            text: "ok",
                            type: "success",
                            timer: 3000,
                            closeOnConfirm: false
                        });
                    }
                    else {
                        swal({
                            title: "Hệ thống",
                            text: "Ôi hỏng.",
                            type: "error",
                            timer: 3000,
                            showConfirmButton: false
                        });
                    }
                },
                error: function (rs) {
                    swal({
                        title: "",
                        text: "Lỗi rồi",
                        type: "error",
                        timer: 2000,
                        showConfirmButton: false
                    });
                },
            });



        } else {
            swal("Cancelled", "Your imaginary file is safe :)", "error");
        }

        window.setTimeout('location.reload()', 1); //Reloads after three seconds
    });
}





function Format_Mony(num) {
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1,")
}



//$("#ID_Form2").serialize() + $("#ID_Form2").serialize();



function Fs_prcDel(_key, _id, _idLang) {
    swal({
        title: "?",
        text: "You will not be able to recover this imaginary file!",
        type: "warning",
        showCancelButton: true,
        confirmButtonColor: "#F00",
        confirmButtonText: "Yes!",
        cancelButtonText: "No",
        closeOnConfirm: false,
        closeOnCancel: false
    }, function (isConfirm) {
        if (isConfirm) {
            $.ajax({
                type: "POST",
                url: "/administrator/Provinces/Fs_prc-del.asp",
                data: {
                    "_key": _key,
                    "lamgG": _idLang,
                    "_id": _id
                },
                cache: false,
                dataType: "html",
                success: function (rs) {

                    alert(rs);
                    if (rs == "1") {
                        swal({
                            title: "Hệ Thống",
                            text: "ok",
                            type: "success",
                            timer: 3000,
                            closeOnConfirm: false
                        });
                    }
                    else {
                        swal({
                            title: "Hệ thống",
                            text: "Ôi hỏng.",
                            type: "error",
                            timer: 3000,
                            showConfirmButton: false
                        });
                    }
                },
                error: function (rs) {
                    swal({
                        title: "",
                        text: "Lỗi rồi",
                        type: "error",
                        timer: 2000,
                        showConfirmButton: false
                    });
                },
            });



        } else {
            swal("Cancelled", "Your imaginary file is safe :)", "error");
        }

        window.setTimeout('location.reload()', 1); //Reloads after three seconds
    });
}