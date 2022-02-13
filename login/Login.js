$('.form').find('input, textarea').on('keyup blur focus', function (e) {

    var $this = $(this),
        label = $this.prev('label');

    if (e.type === 'keyup') {
        if ($this.val() === '') {
            label.removeClass('active highlight');
        } else {
            label.addClass('active highlight');
        }
    } else if (e.type === 'blur') {
        if ($this.val() === '') {
            label.removeClass('active highlight');
        } else {
            label.removeClass('highlight');
        }
    } else if (e.type === 'focus') {

        if ($this.val() === '') {
            label.removeClass('highlight');
        }
        else if ($this.val() !== '') {
            label.addClass('highlight');
        }
    }

});



$('.tab a').on('click', function (e) {

    e.preventDefault();

    $(this).parent().addClass('active');
    $(this).parent().siblings().removeClass('active');

    target = $(this).attr('href');

    $('.tab-content > div').not(target).hide();

    $(target).fadeIn(600);

});

$(document).ready(function () {
    $("#Flogin_adm").submit(function (e) {
        e.preventDefault();

    });

});

function Login_system(str_) {
    $.ajax({
        url: "/administrators/include/ajax_login.asp",
        data: $('#Flogin_adm').serialize() + "&Nid=" + str_,
        type: "POST",
        cache: false,
        dataType: "html",
        success: function (rs) {
            str0 = rs.split(":");
            str1 = str0[0];
            str2 = str0[1]
            if (str1 == "0") {

            }
            else if (str1 == "1") {
                swal({
                    title: str2,
                    type: "success",
                    timer: 2000,
                    showConfirmButton: false
                });
                setTimeout(function () {
                    window.location.replace("/administrators/start.asp");
                }, 2000);

            }
            else if (str1 == "2") {
                swal({
                    title: str2,
                    type: "error",
                    timer: 2000,
                    showConfirmButton: false
                });
            }
            else if (str1 == "3") {
                swal({
                    title: str2,
                    type: "error",
                    timer: 2000,
                    showConfirmButton: false
                });
            }
            else {
                // loi ko kiem soat
            }
        },
        error: function (rs) {
            str0 = rs.split(":");
            str1 = str0[0];
            str2 = str0[1]
            swal({
                title: str2,
                type: "error",
                timer: 2000,
                showConfirmButton: false
            });
        },
    });
}