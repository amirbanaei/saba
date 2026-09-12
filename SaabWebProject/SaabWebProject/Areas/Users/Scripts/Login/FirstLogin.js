
function SendCode(url,urlSendCode) {
    var PhoneNumber = document.getElementById("PhoneNumber");
    var PhoneNumber_Length = PhoneNumber.value.length;

    if (PhoneNumber_Length > 11 || PhoneNumber_Length < 11) {
        PhoneNumber.classList.add("is-Invalid");
        return false;
    }
    $.ajax({
        beforeSend:function (){
            document.getElementById("loading-message").style.display = "block";  
        },
        url: url,
        type: "post",
        dataType: "html",
        data: { phonenumber: PhoneNumber.value },
        success: function (res) {
            document.getElementById("loading-message").style.display = "none";
            console.log(res);
            if (res == "False") {

                document.getElementById("alert_message").innerText = "شماره همراه وارد شده وجود ندارد لطفا با مرکز تماس بگیرید";
                document.getElementById("alert").style.display = "block";

            }
            if (res == "User_NotFound") {
                document.getElementById("alert_message").innerText = "حساب کاربری شما یافت نشد لطفا با مرکز تماس بگیرید";
                document.getElementById("alert").style.display = "block";
                return false;
            }
            if (res == "User_NotActive") {
                document.getElementById("alert_message").innerText = "حساب کاربری شما غیرفعال شده است لطفا با مرکز تماس بگیرید";
                document.getElementById("alert").style.display = "block";
                return false;
            }
            else if(res == "True") {
                window.location.replace(urlSendCode + "?phonenumber=" + PhoneNumber.value);
            }
        },
        error:function (result){
            document.getElementById("loading-message").style.display = "none";
            document.getElementById("alert_message").innerText = "عملیات با خطا مواجه شد لطفا بعدا امتحان فرمائید";
            document.getElementById("alert").style.display = "block";
            console.log(result);
        }
    });
}
