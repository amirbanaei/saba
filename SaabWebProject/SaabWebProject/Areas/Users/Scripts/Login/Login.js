$("#Password").keyup(function () {

    var Password = document.getElementById("Password");
    var Password_Length = Password.value.length;
    if (Password_Length >= 8) {

        if (Password.classList.contains("is-invalid")) {
            Password.classList.remove("is-invalid");
            Password.classList.add("is-valid");
        }
        else {
            Password.classList.add("is-valid");
        }
    }
    else {

        if (Password.classList.contains("is-valid")) {
            Password.classList.remove("is-valid");
            Password.classList.add("is-invalid");

        }
        else {
            Password.classList.add("is-invalid");

        }
        return false;
    }

});

function Login(url,redirect) {
    $("#Login").click(function () {

        var CodeMeli = document.getElementById("usr_NationalCode");
        var Password = document.getElementById("Password");
        var Remember = false;
        if ($("input[name='Remember']").is(':checked')) {
            Remember = true;
        }
        console.log(Remember);
        var PhoneNumber_Length = CodeMeli.value.length;
        var Password_Length = Password.value.length;
        if (PhoneNumber_Length != 10) {
            CodeMeli.classList.add("is-invalid");
            return false;
        } else if (Password_Length < 8) {
            Password.classList.add("is-invalid");
            return false;
        } else {
            console.log("inja");
            var tbuser = new Object();
            tbuser.usr_NationalCode = CodeMeli.value;
            tbuser.usr_Password = Password.value;
            tbuser.usr_RememberMe = Remember;
            $.ajax({
                beforeSend:function (){
                    document.getElementById("loading-message").style.display = "block";
                },
                url: url,
                type: "post",
                dataType: "html",
                data: { login: tbuser },
                success: function (res) {
                    document.getElementById("loading-message").style.display = "none";
                    console.log(res);
                    var result = res.split('-');
                    if (result[0] == "true") {
                        document.location.replace(result[1]);
                        
                    }
                    else {
                        document.getElementById("alert_message").innerText = result[1];
                        document.getElementById("alert").style.display = "block";
                        return false;
                    }

                },error: function (result) {
                    document.getElementById("loading-message").style.display = "none";
                    document.getElementById("alert_message").innerText = "عملیات با خطا مواجه شد لطفا بعدا امتحان کنید";
                    document.getElementById("alert").style.display = "block";
                    console.log(result);
                }
            })
        }

    });
}