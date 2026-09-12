/// صحت سنجی کد تایید چهار رقمی
$(".code").keyup(function () {
    var length = this.value.length;
    if (length == 1) {

        if (this.classList.contains("code_Invalid")) {
            this.classList.remove("code_Invalid");
        }
        this.classList.add("code_valid");

        var nextElement;
        if (this.parentElement.nextElementSibling != null) {
            nextElement = this.parentElement.nextElementSibling.children[0];
            nextElement.focus();
        }
    } else {
        if (this.classList.contains("code_valid")) {
            this.classList.remove("code_valid");
        }
        this.classList.add("code_Invalid");
    }
});

///شمارنده تا ارسال کد بعدی

////

function AccpetCode(urlGet, phonenumber, urlPost) {
    var codes = document.getElementsByClassName("code");
    var arr_code = "";
    for (var i = 0; i < codes.length; i++) {
        var code = codes[i].value;
        arr_code += code;
    }

    if (arr_code.length == 4) {

        $.ajax({
            beforeSend: function () {
                document.getElementById("loading-message").style.display = "block";
            },
            url: urlGet,
            type: "post",
            dataType: "html",
            data: {PhoneNumber: phonenumber, smsCode: arr_code},
            success: function (res) {
                document.getElementById("loading-message").style.display = "none";
                if (res == "false_WrongCode") {
                    document.getElementById("alert_message").innerText = "کد وارد شده صحیح نمی باشد";
                    document.getElementById("alert").style.display = "block";
                } else if (res == "false_Expired") {
                    document.getElementById("alert_message").innerText = "کد وارد شده منقضی شده است لطفا دوباره امتحان کنید !";
                    document.getElementById("alert").style.display = "block";
                } else if (res == "false_NotFoundUser") {
                    document.getElementById("alert_message").innerText = "این شماره همراه ثبت نشده است لطفا با مرکز تماس بگیرید";
                    document.getElementById("alert").style.display = "block";
                } else if (res == "true") {
                    document.location.replace(urlPost);
                }


            },
            error:function (){
                document.getElementById("loading-message").style.display = "none";
                document.getElementById("alert_message").innerText = "عملیات با خطا مواجه شد لطفا بعدا امتحان فرمائید";
                document.getElementById("alert").style.display = "block";
            }

        });
    }
}