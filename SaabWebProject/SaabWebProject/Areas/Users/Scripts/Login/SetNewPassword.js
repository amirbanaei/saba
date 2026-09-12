

var password = document.getElementById("Password");
var confirm_password = document.getElementById("confirm_password");
var Count;
$("#Password").keyup(function () {
    var CountSuccess = 0;
    var lowerCaseLetters = /[a-z]/g;
    var uppercase = /[A-Z]/g;
    var numbers = /[0-9]/g;
    var lower = document.getElementById("lowercase");
    lower.style.display = "block";
    var upper = document.getElementById("uppercase");
    upper.style.display = "block";
    var number = document.getElementById("number");
    number.style.display = "block";
    var character = document.getElementById("character");
    character.style.display = "block";

    /*lowercase*/
    if (password.value.match(lowerCaseLetters)) {

        lower.classList.remove("invalid-feedback");
        lower.classList.add("valid-feedback");
        CountSuccess += 1;

    } else {
        password.classList.remove("is-valid");
        password.classList.add("is-invalid");

        lower.classList.remove("valid-feedback");
        lower.classList.add("invalid-feedback");
    }
    /*lowercase*/

    /*uppercase*/
    if (password.value.match(uppercase)) {

        upper.classList.remove("invalid-feedback");
        upper.classList.add("valid-feedback");
        CountSuccess += 1;

    } else {
        password.classList.remove("is-valid");
        password.classList.add("is-invalid");

        upper.classList.remove("valid-feedback");
        upper.classList.add("invalid-feedback");
    }
    /*uppercase*/

    /*number*/
    if (password.value.match(numbers)) {

        number.classList.remove("invalid-feedback");
        number.classList.add("valid-feedback");
        CountSuccess += 1;

    } else {
        number.classList.remove("is-valid");
        number.classList.add("is-invalid");

        number.classList.remove("valid-feedback");
        number.classList.add("invalid-feedback");
    }
    /*number*/

    /*character*/
    if (password.value.length > 8) {

        character.classList.remove("invalid-feedback");
        character.classList.add("valid-feedback");
        CountSuccess += 1;

    } else {
        character.classList.remove("is-valid");
        character.classList.add("is-invalid");

        character.classList.remove("valid-feedback");
        character.classList.add("invalid-feedback");
    }
    /*character*/
    if (CountSuccess != 4) {
        password.classList.remove("is-valid");
        password.classList.add("is-invalid");
    }
    else {
        password.classList.remove("is-invalid");
        password.classList.add("is-valid");
    }

    Count = CountSuccess;
});

$("#confirm_password").keyup(function () {

    ///match password and confirm password
    if (confirm_password.value != "") {
        console.log(confirm_password.value);
        if (password.value != confirm_password.value) {
            confirm_password.classList.remove("is-valid");
            confirm_password.classList.add("is-invalid");

        }
        else {
            confirm_password.classList.remove("is-invalid");
            confirm_password.classList.add("is-valid");
        }
    }

});

function SetPassword(url,PhoneNumber) {
    if (Count == 4) {
        if (password.value != confirm_password.value) {
            return false;
        }
        $.ajax({
            beforeSend:function (){
                document.getElementById("loading-message").style.display = "block";  
            },
            url: url,
            type: "Post",
            dataType: "html",
            data: { password: password.value, phonenumber: PhoneNumber},
            success: function (res) {
                document.getElementById("loading-message").style.display = "none";
                if (res == "True") {
                    window.location.replace("/Login/Login?message='ثبت نام با موفقیت انجام شد'");
                }
                else if(res == "false") {
                    window.location.replace("/Login/Login?message='کاربر وجود ندارد'");
                }
                else if (res == "False") {
                    window.location.replace("/Login/Login?message='مشکلی پیش اومده لطفا بعدا تلاش فرمائید'");
                }
            },
            error:function (){
                document.getElementById("loading-message").style.display = "none";
                document.getElementById("alert_message").innerText = "عملیات با خطا مواجه شد لطفا بعدا امتحان فرمائید";
                document.getElementById("alert").style.display = "block";  
            }

        })

    }
}
